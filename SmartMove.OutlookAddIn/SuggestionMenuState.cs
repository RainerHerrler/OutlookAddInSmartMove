using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SmartMove.OutlookAddIn
{
    internal sealed class SuggestionMenuState
    {
        private const string SenderSmtpProperty = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";
        private readonly Dictionary<string, MoveSuggestion> slots =
            new Dictionary<string, MoveSuggestion>(StringComparer.OrdinalIgnoreCase);
        private string selectionKey;

        public void Clear()
        {
            selectionKey = null;
            slots.Clear();
        }

        public bool IsVisible(Outlook.Application application, string controlId)
        {
            RefreshIfNeeded(application);
            return slots.ContainsKey(controlId);
        }

        public string GetLabel(Outlook.Application application, string controlId)
        {
            RefreshIfNeeded(application);
            if (!slots.TryGetValue(controlId, out MoveSuggestion suggestion))
            {
                return "Move To …";
            }

            string path = ShortenPath(suggestion.FolderPath);
            if (suggestion.Source == SuggestionSource.Sender)
            {
                return string.Format("Move To {0} — Absender ({1:N0})", path, suggestion.MessageCount);
            }

            string wordLabel = suggestion.MatchedWordCount == 1 ? "1 Signalwort" : suggestion.MatchedWordCount + " Signalwörter";
            return string.Format("Move To {0} — Betreff ({1})", path, wordLabel);
        }

        public int MoveSelection(Outlook.Application application, string controlId)
        {
            RefreshIfNeeded(application);
            if (!slots.TryGetValue(controlId, out MoveSuggestion suggestion))
            {
                throw new InvalidOperationException("Der gewählte SmartMove-Vorschlag ist nicht mehr verfügbar.");
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder destination = null;
            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;
            int movedCount = 0;
            try
            {
                session = application.Session;
                destination = session.GetFolderFromID(suggestion.FolderEntryId, suggestion.StoreId);
                explorer = application.ActiveExplorer();
                selection = explorer?.Selection;
                if (selection == null || selection.Count == 0)
                {
                    throw new InvalidOperationException("Es ist keine E-Mail ausgewählt.");
                }

                // Iterate backwards because moving an item can change Outlook's selection.
                for (int index = selection.Count; index >= 1; index--)
                {
                    object selectedItem = null;
                    object movedItem = null;
                    try
                    {
                        selectedItem = selection[index];
                        if (selectedItem is Outlook.MailItem mailItem)
                        {
                            movedItem = mailItem.Move(destination);
                            movedCount++;
                        }
                    }
                    finally
                    {
                        ReleaseCom(movedItem);
                        ReleaseCom(selectedItem);
                    }
                }

                if (movedCount == 0)
                {
                    throw new InvalidOperationException("Die Auswahl enthält keine verschiebbare E-Mail.");
                }

                SmartMoveAddIn.WriteDiagnostic(
                    "Moved " + movedCount + " item(s) to " + suggestion.FolderPath + " via " + suggestion.Source);
                Clear();
                return movedCount;
            }
            finally
            {
                ReleaseCom(selection);
                ReleaseCom(explorer);
                ReleaseCom(destination);
                ReleaseCom(session);
            }
        }

        private void RefreshIfNeeded(Outlook.Application application)
        {
            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;
            object selectedItem = null;
            try
            {
                explorer = application?.ActiveExplorer();
                selection = explorer?.Selection;
                if (selection == null || selection.Count == 0)
                {
                    Clear();
                    return;
                }

                selectedItem = selection[1];
                if (!(selectedItem is Outlook.MailItem mailItem))
                {
                    Clear();
                    return;
                }

                string currentKey = mailItem.EntryID ?? string.Empty;
                if (!string.IsNullOrEmpty(currentKey) && string.Equals(selectionKey, currentKey, StringComparison.Ordinal))
                {
                    return;
                }

                Clear();
                selectionKey = currentKey;
                string senderAddress = GetSenderAddress(mailItem);
                IReadOnlyCollection<string> words = SubjectWordExtractor.Extract(mailItem.Subject);
                List<MoveSuggestion> senderSuggestions = StatisticsDatabase.GetSenderSuggestions(senderAddress, 2);
                List<MoveSuggestion> wordCandidates = StatisticsDatabase.GetSubjectWordSuggestions(words, 10);

                AddSlots(senderSuggestions, "SmartMove.SenderMove", 2);

                var senderFolderKeys = new HashSet<string>(
                    senderSuggestions.Select(suggestion => suggestion.FolderKey),
                    StringComparer.OrdinalIgnoreCase);
                List<MoveSuggestion> distinctWordSuggestions = wordCandidates
                    .Where(suggestion => !senderFolderKeys.Contains(suggestion.FolderKey))
                    .Take(2)
                    .ToList();
                if (distinctWordSuggestions.Count < 2)
                {
                    foreach (MoveSuggestion candidate in wordCandidates)
                    {
                        if (distinctWordSuggestions.Any(existing =>
                            string.Equals(existing.FolderKey, candidate.FolderKey, StringComparison.OrdinalIgnoreCase)))
                        {
                            continue;
                        }
                        distinctWordSuggestions.Add(candidate);
                        if (distinctWordSuggestions.Count == 2)
                        {
                            break;
                        }
                    }
                }
                AddSlots(distinctWordSuggestions, "SmartMove.WordMove", 2);
            }
            catch (Exception exception)
            {
                SmartMoveAddIn.WriteDiagnostic("Suggestion refresh failed: " + exception);
                Clear();
            }
            finally
            {
                ReleaseCom(selectedItem);
                ReleaseCom(selection);
                ReleaseCom(explorer);
            }
        }

        private void AddSlots(IEnumerable<MoveSuggestion> suggestions, string idPrefix, int limit)
        {
            int index = 1;
            foreach (MoveSuggestion suggestion in suggestions.Take(limit))
            {
                slots[idPrefix + index] = suggestion;
                index++;
            }
        }

        private static string GetSenderAddress(Outlook.MailItem mailItem)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = mailItem.PropertyAccessor;
                object value = accessor.GetProperty(SenderSmtpProperty);
                string smtp = value as string;
                if (!string.IsNullOrWhiteSpace(smtp))
                {
                    return smtp.Trim().ToLowerInvariant();
                }
            }
            catch
            {
                // Fall back to Outlook's regular sender address below.
            }
            finally
            {
                ReleaseCom(accessor);
            }

            return string.IsNullOrWhiteSpace(mailItem.SenderEmailAddress)
                ? string.Empty
                : mailItem.SenderEmailAddress.Trim().ToLowerInvariant();
        }

        private static string ShortenPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            string[] levels = path.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (levels.Length <= 2)
            {
                return string.Join("\\", levels);
            }

            return levels[levels.Length - 2] + "\\" + levels[levels.Length - 1];
        }

        private static void ReleaseCom(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
