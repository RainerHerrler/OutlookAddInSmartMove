using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SmartMove.OutlookAddIn
{
    [ComVisible(true)]
    [Guid("62A782A4-0AC8-4B8E-BC92-E5D66D1360A8")]
    [ProgId("SmartMove.OutlookAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class SmartMoveAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private const string ExplorerRibbonId = "Microsoft.Outlook.Explorer";
        private object outlookApplication;
        private OutlookStatisticsScanner activeScanner;
        private readonly SuggestionMenuState suggestionMenu = new SuggestionMenuState();

        public string GetCustomUI(string ribbonId)
        {
            WriteDiagnostic("GetCustomUI: " + (ribbonId ?? "<null>"));

            if (!string.Equals(ribbonId, ExplorerRibbonId, StringComparison.Ordinal))
            {
                return null;
            }

            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <contextMenus>
    <contextMenu idMso=""ContextMenuMailItem"">
      <menu id=""SmartMove.ContextMenu""
            label=""SmartMove""
            screentip=""E-Mails intelligent ablegen""
            supertip=""Erstellt und exportiert die lokale SmartMove-Statistik."">
        <button id=""SmartMove.SenderMove1""
                getLabel=""GetSuggestionLabel""
                getVisible=""GetSuggestionVisible""
                onAction=""OnMoveSuggestion"" />
        <button id=""SmartMove.SenderMove2""
                getLabel=""GetSuggestionLabel""
                getVisible=""GetSuggestionVisible""
                onAction=""OnMoveSuggestion"" />
        <button id=""SmartMove.WordMove1""
                getLabel=""GetSuggestionLabel""
                getVisible=""GetSuggestionVisible""
                onAction=""OnMoveSuggestion"" />
        <button id=""SmartMove.WordMove2""
                getLabel=""GetSuggestionLabel""
                getVisible=""GetSuggestionVisible""
                onAction=""OnMoveSuggestion"" />
        <menuSeparator id=""SmartMove.SuggestionSeparator"" />
        <button id=""SmartMove.SearchLinkedInButton""
                label=""Absender auf LinkedIn suchen""
                screentip=""LinkedIn-Personensuche öffnen""
                supertip=""Sucht den Namen des Absenders im Standardbrowser auf LinkedIn.""
                onAction=""OnSearchLinkedIn"" />
        <menuSeparator id=""SmartMove.ToolsSeparator"" />
        <button id=""SmartMove.InitialScanButton""
                label=""Statistik initial erstellen""
                screentip=""Ablageordner manuell auswerten""
                onAction=""OnInitialScan"" />
        <button id=""SmartMove.ExportButton""
                label=""Statistik als JSON exportieren""
                screentip=""Statistik nach Häufigkeit sortiert exportieren""
                onAction=""OnExportStatistics"" />
      </menu>
    </contextMenu>
  </contextMenus>
</customUI>";
        }

        // Ribbon callbacks must be public so Office can invoke them through COM.
        public bool GetSuggestionVisible(Office.IRibbonControl control)
        {
            Outlook.Application application = outlookApplication as Outlook.Application;
            return application != null && suggestionMenu.IsVisible(application, control.Id);
        }

        public string GetSuggestionLabel(Office.IRibbonControl control)
        {
            Outlook.Application application = outlookApplication as Outlook.Application;
            return application == null ? "Move To …" : suggestionMenu.GetLabel(application, control.Id);
        }

        public void OnMoveSuggestion(Office.IRibbonControl control)
        {
            WriteDiagnostic("OnMoveSuggestion: " + control.Id);
            try
            {
                Outlook.Application application = outlookApplication as Outlook.Application;
                if (application == null)
                {
                    throw new InvalidOperationException("Die Outlook-Anwendung ist nicht verfügbar.");
                }
                suggestionMenu.MoveSelection(application, control.Id);
            }
            catch (Exception exception)
            {
                WriteDiagnostic("Move suggestion failed: " + exception);
                MessageBox.Show(
                    "Die E-Mail konnte nicht verschoben werden.\n\n" + exception.Message,
                    "SmartMove",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void OnSearchLinkedIn(Office.IRibbonControl control)
        {
            WriteDiagnostic("OnSearchLinkedIn");
            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;
            object selectedItem = null;
            try
            {
                Outlook.Application application = outlookApplication as Outlook.Application;
                if (application == null)
                {
                    throw new InvalidOperationException("Die Outlook-Anwendung ist nicht verfügbar.");
                }

                explorer = application.ActiveExplorer();
                selection = explorer?.Selection;
                if (selection == null || selection.Count == 0)
                {
                    throw new InvalidOperationException("Es ist keine E-Mail ausgewählt.");
                }

                selectedItem = selection[1];
                if (!(selectedItem is Outlook.MailItem mailItem))
                {
                    throw new InvalidOperationException("Die Auswahl enthält keine E-Mail.");
                }

                string senderName = mailItem.SenderName?.Trim();
                if (string.IsNullOrWhiteSpace(senderName))
                {
                    throw new InvalidOperationException("Für diese E-Mail ist kein Absendername verfügbar.");
                }

                string url = "https://www.linkedin.com/search/results/people/?keywords=" +
                    Uri.EscapeDataString(senderName);
                using (Process process = Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }))
                {
                }
            }
            catch (Exception exception)
            {
                WriteDiagnostic("LinkedIn search failed: " + exception.GetType().FullName);
                MessageBox.Show(
                    "Die LinkedIn-Suche konnte nicht geöffnet werden.\n\n" + exception.Message,
                    "SmartMove",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                ReleaseCom(selectedItem);
                ReleaseCom(selection);
                ReleaseCom(explorer);
            }
        }

        public void OnInitialScan(Office.IRibbonControl control)
        {
            WriteDiagnostic("OnInitialScan");
            if (activeScanner != null)
            {
                MessageBox.Show("Ein Statistik-Scan läuft bereits.", "SmartMove", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult answer = MessageBox.Show(
                "SmartMove wertet jetzt alle E-Mail-Ablageordner aus.\n\n" +
                "Posteingang, Gesendet, Entwürfe, Papierkorb, Junk und technische Ordner werden ausgelassen. " +
                "Eine vorhandene Statistik wird ersetzt.\n\nScan jetzt starten?",
                "SmartMove – Statistik initial erstellen",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (answer != DialogResult.Yes)
            {
                return;
            }

            var application = outlookApplication as Outlook.Application;
            if (application == null)
            {
                MessageBox.Show("Die Outlook-Anwendung ist nicht verfügbar.", "SmartMove", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            activeScanner = new OutlookStatisticsScanner(application, () =>
            {
                activeScanner = null;
                suggestionMenu.Clear();
            });
            activeScanner.Start();
        }

        public void OnExportStatistics(Office.IRibbonControl control)
        {
            WriteDiagnostic("OnExportStatistics");
            if (activeScanner != null)
            {
                MessageBox.Show("Bitte warten Sie, bis der laufende Scan beendet ist.", "SmartMove", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                using (var dialog = new SaveFileDialog())
                {
                    dialog.Title = "SmartMove-Statistik exportieren";
                    dialog.Filter = "JSON-Datei (*.json)|*.json";
                    dialog.DefaultExt = "json";
                    dialog.AddExtension = true;
                    dialog.FileName = "SmartMove-statistics-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".json";
                    dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    if (dialog.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    int senderCount = StatisticsDatabase.ExportJson(dialog.FileName);
                    MessageBox.Show(
                        string.Format("Die Statistik für {0:N0} Absender wurde exportiert:\n\n{1}", senderCount, dialog.FileName),
                        "SmartMove",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception exception)
            {
                WriteDiagnostic("Export failed: " + exception);
                MessageBox.Show("Der Export ist fehlgeschlagen.\n\n" + exception.Message, "SmartMove", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInstance, ref Array custom)
        {
            outlookApplication = application;
            WriteDiagnostic("OnConnection: " + connectMode);
            try
            {
                SmartMoveConfiguration.EnsureFileExists();
            }
            catch (Exception exception)
            {
                WriteDiagnostic("Configuration initialization failed: " + exception);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            ReleaseOutlookApplication();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            ReleaseOutlookApplication();
        }

        private void ReleaseOutlookApplication()
        {
            activeScanner?.Dispose();
            activeScanner = null;
            suggestionMenu.Clear();
            if (outlookApplication != null && Marshal.IsComObject(outlookApplication))
            {
                Marshal.FinalReleaseComObject(outlookApplication);
            }

            outlookApplication = null;
        }

        private static void ReleaseCom(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }

        internal static void WriteDiagnostic(string message)
        {
            try
            {
                string directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "SmartMove");
                Directory.CreateDirectory(directory);
                File.AppendAllText(
                    Path.Combine(directory, "SmartMove.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + message + Environment.NewLine);
            }
            catch
            {
                // Diagnostics must never prevent Outlook from loading the add-in.
            }
        }
    }
}
