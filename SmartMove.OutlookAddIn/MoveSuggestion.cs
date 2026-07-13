namespace SmartMove.OutlookAddIn
{
    internal enum SuggestionSource
    {
        Sender,
        SubjectWords
    }

    internal sealed class MoveSuggestion
    {
        public string StoreId { get; set; }
        public string FolderEntryId { get; set; }
        public string FolderPath { get; set; }
        public SuggestionSource Source { get; set; }
        public long MessageCount { get; set; }
        public int MatchedWordCount { get; set; }
        public double Score { get; set; }

        public string FolderKey => StoreId + "\n" + FolderEntryId;
    }
}
