using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SmartMove.OutlookAddIn
{
    internal sealed class OutlookStatisticsScanner : IDisposable
    {
        private const int MaxRowsPerTick = 100;
        private const int MaxWorkMillisecondsPerTick = 25;
        private const int PauseBetweenBatchesMilliseconds = 150;
        private const string MessageClassColumn = "http://schemas.microsoft.com/mapi/proptag/0x001A001F";
        private const string SenderSmtpColumn = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";
        private const string SenderAddressColumn = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F";
        private const string SubjectColumn = "http://schemas.microsoft.com/mapi/proptag/0x0037001F";
        private const string FolderTypeProperty = "http://schemas.microsoft.com/mapi/proptag/0x36010003";

        private readonly Outlook.Application application;
        private readonly Action finished;
        private readonly List<FolderDescriptor> folders = new List<FolderDescriptor>();
        private readonly Timer timer;
        private Outlook.NameSpace session;
        private StatisticsDatabase database;
        private ScanProgressForm progressForm;
        private Outlook.MAPIFolder currentFolder;
        private Outlook.Table currentTable;
        private int folderIndex;
        private long currentFolderId;
        private long inspectedRows;
        private long matchedEmails;
        private bool cancelRequested;
        private bool disposed;

        public OutlookStatisticsScanner(Outlook.Application application, Action finished)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.finished = finished;
            timer = new Timer { Interval = PauseBetweenBatchesMilliseconds };
            timer.Tick += ProcessNextBatch;
        }

        public void Start()
        {
            try
            {
                session = application.Session;
                DiscoverFolders();
                if (folders.Count == 0)
                {
                    throw new InvalidOperationException("Es wurden keine auswertbaren E-Mail-Ablageordner gefunden.");
                }

                database = new StatisticsDatabase();
                database.BeginInitialScan();
                progressForm = new ScanProgressForm();
                progressForm.CancelRequested += (sender, args) => cancelRequested = true;
                progressForm.Show();
                timer.Start();
            }
            catch (Exception exception)
            {
                SmartMoveAddIn.WriteDiagnostic("Scan start failed: " + exception);
                MessageBox.Show(
                    "Die Statistik konnte nicht gestartet werden.\n\n" + exception.Message,
                    "SmartMove",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Dispose();
                finished?.Invoke();
            }
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }
            disposed = true;

            timer.Stop();
            timer.Dispose();
            ReleaseCurrentFolder();
            database?.Dispose();
            database = null;
            ReleaseCom(session);
            session = null;
            if (progressForm != null && !progressForm.IsDisposed)
            {
                progressForm.CloseAfterScan();
            }
            progressForm?.Dispose();
            progressForm = null;
        }

        private void ProcessNextBatch(object sender, EventArgs args)
        {
            timer.Stop();
            try
            {
                if (cancelRequested)
                {
                    database.Cancel(matchedEmails, folderIndex);
                    Finish("Der Scan wurde abgebrochen. Die bisherige Teilstatistik bleibt gespeichert.", MessageBoxIcon.Information);
                    return;
                }

                int remaining = MaxRowsPerTick;
                var workTime = Stopwatch.StartNew();
                while (remaining > 0 && workTime.ElapsedMilliseconds < MaxWorkMillisecondsPerTick)
                {
                    if (currentTable == null && !OpenNextFolder())
                    {
                        database.Complete(matchedEmails, folders.Count);
                        Finish(
                            string.Format(
                                "Die Statistik ist fertig.\n\n{0:N0} E-Mails aus {1:N0} Ordnern wurden gezählt.",
                                matchedEmails,
                                folders.Count),
                            MessageBoxIcon.Information);
                        return;
                    }

                    if (currentTable.EndOfTable)
                    {
                        ReleaseCurrentFolder();
                        continue;
                    }

                    Outlook.Row row = null;
                    try
                    {
                        row = currentTable.GetNextRow();
                        inspectedRows++;
                        string messageClass = ReadString(row, MessageClassColumn);
                        if (messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase))
                        {
                            matchedEmails++;
                            string senderAddress = NormalizeAddress(ReadString(row, SenderSmtpColumn));
                            if (string.IsNullOrEmpty(senderAddress))
                            {
                                senderAddress = NormalizeAddress(ReadString(row, SenderAddressColumn));
                            }

                            if (!string.IsNullOrEmpty(senderAddress))
                            {
                                database.Increment(senderAddress, currentFolderId);
                            }

                            string subject = ReadString(row, SubjectColumn);
                            foreach (string word in SubjectWordExtractor.Extract(subject))
                            {
                                database.IncrementSubjectWord(word, currentFolderId);
                            }
                        }
                    }
                    finally
                    {
                        ReleaseCom(row);
                    }

                    remaining--;
                }

                FolderDescriptor descriptor = folders[Math.Max(0, folderIndex - 1)];
                progressForm.SetProgress(descriptor.Path, folderIndex - 1, folders.Count, inspectedRows, matchedEmails);
                timer.Start();
            }
            catch (Exception exception)
            {
                SmartMoveAddIn.WriteDiagnostic("Scan failed: " + exception);
                try
                {
                    database?.Cancel(matchedEmails, folderIndex);
                }
                catch (Exception databaseException)
                {
                    SmartMoveAddIn.WriteDiagnostic("Scan cleanup failed: " + databaseException);
                }
                Finish("Der Scan wurde wegen eines Fehlers beendet.\n\n" + exception.Message, MessageBoxIcon.Error);
            }
        }

        private bool OpenNextFolder()
        {
            while (folderIndex < folders.Count)
            {
                FolderDescriptor descriptor = folders[folderIndex++];
                try
                {
                    currentFolder = session.GetFolderFromID(descriptor.EntryId, descriptor.StoreId);
                    currentFolderId = database.EnsureFolder(descriptor.StoreId, descriptor.EntryId, descriptor.Path);
                    currentTable = currentFolder.GetTable(null, Outlook.OlTableContents.olUserItems);
                    currentTable.Columns.RemoveAll();
                    currentTable.Columns.Add(MessageClassColumn);
                    currentTable.Columns.Add(SenderSmtpColumn);
                    currentTable.Columns.Add(SenderAddressColumn);
                    currentTable.Columns.Add(SubjectColumn);
                    progressForm.SetProgress(descriptor.Path, folderIndex - 1, folders.Count, inspectedRows, matchedEmails);
                    return true;
                }
                catch (Exception exception)
                {
                    SmartMoveAddIn.WriteDiagnostic("Skipping folder " + descriptor.Path + ": " + exception.Message);
                    ReleaseCurrentFolder();
                }
            }

            return false;
        }

        private void DiscoverFolders()
        {
            Outlook.Stores stores = null;
            try
            {
                stores = session.Stores;
                for (int index = 1; index <= stores.Count; index++)
                {
                    Outlook.Store store = null;
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        store = stores[index];
                        root = store.GetRootFolder();
                        var excludedItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        var excludedTrees = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderInbox, excludedItems);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderSentMail, excludedItems);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderOutbox, excludedItems);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderDrafts, excludedItems);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderDeletedItems, excludedTrees);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderJunk, excludedTrees);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderConflicts, excludedTrees);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderSyncIssues, excludedTrees);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderLocalFailures, excludedTrees);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderServerFailures, excludedTrees);
                        AddDefaultFolder(store, Outlook.OlDefaultFolders.olFolderRssFeeds, excludedTrees);
                        excludedItems.Add(root.EntryID);
                        DiscoverFolder(root, excludedItems, excludedTrees, false);
                    }
                    catch (Exception exception)
                    {
                        SmartMoveAddIn.WriteDiagnostic("Skipping store: " + exception.Message);
                    }
                    finally
                    {
                        ReleaseCom(root);
                        ReleaseCom(store);
                    }
                }
            }
            finally
            {
                ReleaseCom(stores);
            }
        }

        private void DiscoverFolder(
            Outlook.MAPIFolder folder,
            HashSet<string> excludedItems,
            HashSet<string> excludedTrees,
            bool ancestorExcluded)
        {
            string entryId = folder.EntryID;
            bool treeExcluded = ancestorExcluded || excludedTrees.Contains(entryId) || IsSearchFolder(folder);
            if (!treeExcluded &&
                !excludedItems.Contains(entryId) &&
                folder.DefaultItemType == Outlook.OlItemType.olMailItem)
            {
                folders.Add(new FolderDescriptor(folder.StoreID, entryId, folder.FolderPath));
            }

            if (treeExcluded)
            {
                return;
            }

            Outlook.Folders children = null;
            try
            {
                children = folder.Folders;
                for (int index = 1; index <= children.Count; index++)
                {
                    Outlook.MAPIFolder child = null;
                    try
                    {
                        child = children[index];
                        DiscoverFolder(child, excludedItems, excludedTrees, false);
                    }
                    finally
                    {
                        ReleaseCom(child);
                    }
                }
            }
            finally
            {
                ReleaseCom(children);
            }
        }

        private static void AddDefaultFolder(Outlook.Store store, Outlook.OlDefaultFolders folderType, HashSet<string> target)
        {
            Outlook.MAPIFolder folder = null;
            try
            {
                folder = store.GetDefaultFolder(folderType);
                if (folder != null && !string.IsNullOrEmpty(folder.EntryID))
                {
                    target.Add(folder.EntryID);
                }
            }
            catch
            {
                // Not every store supports every default folder.
            }
            finally
            {
                ReleaseCom(folder);
            }
        }

        private static bool IsSearchFolder(Outlook.MAPIFolder folder)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = folder.PropertyAccessor;
                object value = accessor.GetProperty(FolderTypeProperty);
                return Convert.ToInt32(value) == 2;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(accessor);
            }
        }

        private static string ReadString(Outlook.Row row, string column)
        {
            try
            {
                object value = row[column];
                return value as string ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string NormalizeAddress(string address)
        {
            return string.IsNullOrWhiteSpace(address) ? string.Empty : address.Trim().ToLowerInvariant();
        }

        private void ReleaseCurrentFolder()
        {
            ReleaseCom(currentTable);
            currentTable = null;
            ReleaseCom(currentFolder);
            currentFolder = null;
            currentFolderId = 0;
        }

        private void Finish(string message, MessageBoxIcon icon)
        {
            Dispose();
            MessageBox.Show(message, "SmartMove", MessageBoxButtons.OK, icon);
            finished?.Invoke();
        }

        private static void ReleaseCom(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }

        private sealed class FolderDescriptor
        {
            public FolderDescriptor(string storeId, string entryId, string path)
            {
                StoreId = storeId;
                EntryId = entryId;
                Path = path;
            }

            public string StoreId { get; }
            public string EntryId { get; }
            public string Path { get; }
        }
    }
}
