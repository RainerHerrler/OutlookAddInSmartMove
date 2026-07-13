using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace SmartMove.OutlookAddIn
{
    internal sealed class StatisticsDatabase : IDisposable
    {
        private const int CommitInterval = 1000;
        private readonly SQLiteConnection connection;
        private SQLiteTransaction transaction;
        private SQLiteCommand incrementCommand;
        private SQLiteCommand incrementWordCommand;
        private int changesSinceCommit;

        public static string DataDirectory => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SmartMove");

        public static string DatabasePath => Path.Combine(DataDirectory, "smartmove.db");

        public StatisticsDatabase(string databasePath = null)
        {
            string path = databasePath ?? DatabasePath;
            string directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }
            connection = new SQLiteConnection("Data Source=" + path + ";Version=3;Foreign Keys=True;");
            connection.Open();
            ExecuteNonQuery("PRAGMA journal_mode=WAL;");
            ExecuteNonQuery("PRAGMA synchronous=NORMAL;");
            CreateSchema();
        }

        public void BeginInitialScan()
        {
            using (SQLiteTransaction reset = connection.BeginTransaction())
            {
                ExecuteNonQuery("DELETE FROM subject_word_folder_stats;", reset);
                ExecuteNonQuery("DELETE FROM sender_folder_stats;", reset);
                ExecuteNonQuery("DELETE FROM folders;", reset);
                SetMetadata("scan_status", "running", reset);
                SetMetadata("scan_started_utc", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture), reset);
                reset.Commit();
            }

            BeginTransaction();
        }

        public long EnsureFolder(string storeId, string entryId, string folderPath)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = @"
INSERT INTO folders(store_id, entry_id, folder_path)
VALUES(@store_id, @entry_id, @folder_path)
ON CONFLICT(store_id, entry_id) DO UPDATE SET folder_path = excluded.folder_path;
SELECT id FROM folders WHERE store_id = @store_id AND entry_id = @entry_id;";
                command.Parameters.AddWithValue("@store_id", storeId);
                command.Parameters.AddWithValue("@entry_id", entryId);
                command.Parameters.AddWithValue("@folder_path", folderPath);
                return Convert.ToInt64(command.ExecuteScalar(), CultureInfo.InvariantCulture);
            }
        }

        public void Increment(string senderAddress, long folderId)
        {
            EnsureIncrementCommand();
            incrementCommand.Parameters["@sender"].Value = senderAddress;
            incrementCommand.Parameters["@folder_id"].Value = folderId;
            incrementCommand.Parameters["@seen_utc"].Value = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
            incrementCommand.ExecuteNonQuery();

            RegisterChange();
        }

        public void IncrementSubjectWord(string word, long folderId)
        {
            EnsureIncrementWordCommand();
            incrementWordCommand.Parameters["@word"].Value = word;
            incrementWordCommand.Parameters["@folder_id"].Value = folderId;
            incrementWordCommand.Parameters["@seen_utc"].Value = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
            incrementWordCommand.ExecuteNonQuery();
            RegisterChange();
        }

        public void Complete(long scannedMessages, int scannedFolders)
        {
            CommitCurrentTransaction();
            using (SQLiteTransaction final = connection.BeginTransaction())
            {
                SetMetadata("scan_status", "complete", final);
                SetMetadata("last_scan_utc", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture), final);
                SetMetadata("last_scan_message_count", scannedMessages.ToString(CultureInfo.InvariantCulture), final);
                SetMetadata("last_scan_folder_count", scannedFolders.ToString(CultureInfo.InvariantCulture), final);
                final.Commit();
            }
        }

        public void Cancel(long scannedMessages, int scannedFolders)
        {
            CommitCurrentTransaction();
            using (SQLiteTransaction final = connection.BeginTransaction())
            {
                SetMetadata("scan_status", "cancelled", final);
                SetMetadata("last_scan_message_count", scannedMessages.ToString(CultureInfo.InvariantCulture), final);
                SetMetadata("last_scan_folder_count", scannedFolders.ToString(CultureInfo.InvariantCulture), final);
                final.Commit();
            }
        }

        public static int ExportJson(string destinationPath, string sourceDatabasePath = null)
        {
            string databasePath = sourceDatabasePath ?? DatabasePath;
            if (!File.Exists(databasePath))
            {
                throw new InvalidOperationException("Es ist noch keine SmartMove-Statistik vorhanden.");
            }

            // Apply non-destructive schema migrations before opening read-only.
            using (var migration = new StatisticsDatabase(databasePath))
            {
            }

            var senders = new System.Collections.Generic.List<SenderExport>();
            SenderExport currentSender = null;
            var words = new System.Collections.Generic.List<WordExport>();
            WordExport currentWord = null;
            string lastScanUtc = null;

            using (var connection = new SQLiteConnection("Data Source=" + databasePath + ";Version=3;Read Only=True;"))
            {
                connection.Open();
                using (SQLiteCommand metadata = connection.CreateCommand())
                {
                    metadata.CommandText = "SELECT value FROM metadata WHERE key = 'last_scan_utc';";
                    object value = metadata.ExecuteScalar();
                    if (value != null && value != DBNull.Value)
                    {
                        lastScanUtc = Convert.ToString(value, CultureInfo.InvariantCulture);
                    }
                }

                using (SQLiteCommand command = connection.CreateCommand())
                {
                    command.CommandText = @"
WITH sender_totals AS (
    SELECT sender_address, SUM(message_count) AS total_count
    FROM sender_folder_stats
    GROUP BY sender_address
)
SELECT s.sender_address, f.folder_path, s.message_count
FROM sender_folder_stats s
JOIN folders f ON f.id = s.folder_id
JOIN sender_totals t ON t.sender_address = s.sender_address
ORDER BY t.total_count DESC, s.sender_address COLLATE NOCASE,
         s.message_count DESC, f.folder_path COLLATE NOCASE;";
                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string sender = reader.GetString(0);
                            string folder = reader.GetString(1);
                            long count = reader.GetInt64(2);
                            if (currentSender == null ||
                                !string.Equals(currentSender.Address, sender, StringComparison.OrdinalIgnoreCase))
                            {
                                currentSender = new SenderExport(sender);
                                senders.Add(currentSender);
                            }

                            currentSender.Folders.Add(new FolderExport(folder, count));
                        }
                    }
                }

                using (SQLiteCommand command = connection.CreateCommand())
                {
                    command.CommandText = @"
WITH word_totals AS (
    SELECT word, SUM(message_count) AS total_count
    FROM subject_word_folder_stats
    GROUP BY word
)
SELECT s.word, f.folder_path, s.message_count
FROM subject_word_folder_stats s
JOIN folders f ON f.id = s.folder_id
JOIN word_totals t ON t.word = s.word
ORDER BY t.total_count DESC, s.word COLLATE NOCASE,
         s.message_count DESC, f.folder_path COLLATE NOCASE;";
                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string word = reader.GetString(0);
                            string folder = reader.GetString(1);
                            long count = reader.GetInt64(2);
                            if (currentWord == null ||
                                !string.Equals(currentWord.Word, word, StringComparison.OrdinalIgnoreCase))
                            {
                                currentWord = new WordExport(word);
                                words.Add(currentWord);
                            }

                            currentWord.Folders.Add(new FolderExport(folder, count));
                        }
                    }
                }
            }

            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var json = new StringBuilder();
            json.AppendLine("{");
            json.AppendLine("  \"version\": 2,");
            json.Append("  \"lastScanUtc\": ").Append(serializer.Serialize(lastScanUtc)).AppendLine(",");
            json.AppendLine("  \"senders\": {");
            for (int senderIndex = 0; senderIndex < senders.Count; senderIndex++)
            {
                SenderExport sender = senders[senderIndex];
                json.Append("    ").Append(serializer.Serialize(sender.Address)).AppendLine(": {");
                for (int folderIndex = 0; folderIndex < sender.Folders.Count; folderIndex++)
                {
                    FolderExport folder = sender.Folders[folderIndex];
                    json.Append("      ")
                        .Append(serializer.Serialize(folder.Path))
                        .Append(": ")
                        .Append(folder.Count.ToString(CultureInfo.InvariantCulture));
                    if (folderIndex < sender.Folders.Count - 1)
                    {
                        json.Append(',');
                    }
                    json.AppendLine();
                }
                json.Append("    }");
                if (senderIndex < senders.Count - 1)
                {
                    json.Append(',');
                }
                json.AppendLine();
            }
            json.AppendLine("  },");
            json.AppendLine("  \"subjectWords\": {");
            for (int wordIndex = 0; wordIndex < words.Count; wordIndex++)
            {
                WordExport word = words[wordIndex];
                json.Append("    ").Append(serializer.Serialize(word.Word)).AppendLine(": {");
                for (int folderIndex = 0; folderIndex < word.Folders.Count; folderIndex++)
                {
                    FolderExport folder = word.Folders[folderIndex];
                    json.Append("      ")
                        .Append(serializer.Serialize(folder.Path))
                        .Append(": ")
                        .Append(folder.Count.ToString(CultureInfo.InvariantCulture));
                    if (folderIndex < word.Folders.Count - 1)
                    {
                        json.Append(',');
                    }
                    json.AppendLine();
                }
                json.Append("    }");
                if (wordIndex < words.Count - 1)
                {
                    json.Append(',');
                }
                json.AppendLine();
            }
            json.AppendLine("  }");
            json.AppendLine("}");
            File.WriteAllText(destinationPath, json.ToString(), new UTF8Encoding(false));
            return senders.Count;
        }

        public static List<MoveSuggestion> GetSenderSuggestions(string senderAddress, int limit, string sourceDatabasePath = null)
        {
            var result = new List<MoveSuggestion>();
            string databasePath = sourceDatabasePath ?? DatabasePath;
            if (string.IsNullOrWhiteSpace(senderAddress) || limit <= 0 || !File.Exists(databasePath))
            {
                return result;
            }

            try
            {
                using (var connection = new SQLiteConnection("Data Source=" + databasePath + ";Version=3;Read Only=True;"))
                {
                    connection.Open();
                    using (SQLiteCommand command = connection.CreateCommand())
                    {
                        command.CommandText = @"
SELECT f.store_id, f.entry_id, f.folder_path, s.message_count
FROM sender_folder_stats s
JOIN folders f ON f.id = s.folder_id
WHERE s.sender_address = @sender COLLATE NOCASE
ORDER BY s.message_count DESC, f.folder_path COLLATE NOCASE
LIMIT @limit;";
                        command.Parameters.AddWithValue("@sender", senderAddress);
                        command.Parameters.AddWithValue("@limit", limit);
                        using (SQLiteDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                long count = reader.GetInt64(3);
                                result.Add(new MoveSuggestion
                                {
                                    StoreId = reader.GetString(0),
                                    FolderEntryId = reader.GetString(1),
                                    FolderPath = reader.GetString(2),
                                    Source = SuggestionSource.Sender,
                                    MessageCount = count,
                                    Score = count
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                SmartMoveAddIn.WriteDiagnostic("Sender suggestion query failed: " + exception.Message);
            }

            return result;
        }

        public static List<MoveSuggestion> GetSubjectWordSuggestions(
            IEnumerable<string> subjectWords,
            int limit,
            string sourceDatabasePath = null)
        {
            var words = subjectWords?
                .Where(word => !string.IsNullOrWhiteSpace(word))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();
            string databasePath = sourceDatabasePath ?? DatabasePath;
            if (words.Count == 0 || limit <= 0 || !File.Exists(databasePath))
            {
                return new List<MoveSuggestion>();
            }

            var byFolder = new Dictionary<string, MoveSuggestion>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var connection = new SQLiteConnection("Data Source=" + databasePath + ";Version=3;Read Only=True;"))
                {
                    connection.Open();
                    using (SQLiteCommand command = connection.CreateCommand())
                    {
                        var parameterNames = new List<string>();
                        for (int index = 0; index < words.Count; index++)
                        {
                            string parameterName = "@word" + index.ToString(CultureInfo.InvariantCulture);
                            parameterNames.Add(parameterName);
                            command.Parameters.AddWithValue(parameterName, words[index]);
                        }

                        command.CommandText = @"
WITH word_totals AS (
    SELECT word, SUM(message_count) AS total_count
    FROM subject_word_folder_stats
    WHERE word IN (" + string.Join(",", parameterNames) + @")
    GROUP BY word
)
SELECT f.store_id, f.entry_id, f.folder_path,
       s.word, s.message_count, t.total_count
FROM subject_word_folder_stats s
JOIN folders f ON f.id = s.folder_id
JOIN word_totals t ON t.word = s.word
WHERE s.word IN (" + string.Join(",", parameterNames) + ");";

                        using (SQLiteDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string storeId = reader.GetString(0);
                                string entryId = reader.GetString(1);
                                string key = storeId + "\n" + entryId;
                                long count = reader.GetInt64(4);
                                long totalCount = reader.GetInt64(5);
                                if (!byFolder.TryGetValue(key, out MoveSuggestion suggestion))
                                {
                                    suggestion = new MoveSuggestion
                                    {
                                        StoreId = storeId,
                                        FolderEntryId = entryId,
                                        FolderPath = reader.GetString(2),
                                        Source = SuggestionSource.SubjectWords
                                    };
                                    byFolder.Add(key, suggestion);
                                }

                                // A frequent word/folder relation contributes evidence. The
                                // final multiplier below rewards several independent words
                                // pointing to the same destination.
                                double probability = totalCount == 0 ? 0 : (double)count / totalCount;
                                suggestion.Score += probability * Math.Log(1.0 + count);
                                suggestion.MessageCount += count;
                                suggestion.MatchedWordCount++;
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                SmartMoveAddIn.WriteDiagnostic("Subject-word suggestion query failed: " + exception.Message);
                return new List<MoveSuggestion>();
            }

            foreach (MoveSuggestion suggestion in byFolder.Values)
            {
                suggestion.Score *= 1.0 + (0.5 * Math.Max(0, suggestion.MatchedWordCount - 1));
            }

            return byFolder.Values
                .OrderByDescending(suggestion => suggestion.Score)
                .ThenByDescending(suggestion => suggestion.MatchedWordCount)
                .ThenByDescending(suggestion => suggestion.MessageCount)
                .ThenBy(suggestion => suggestion.FolderPath, StringComparer.OrdinalIgnoreCase)
                .Take(limit)
                .ToList();
        }

        public void Dispose()
        {
            CommitCurrentTransaction();
            incrementCommand?.Dispose();
            incrementWordCommand?.Dispose();
            connection.Dispose();
        }

        private void CreateSchema()
        {
            ExecuteNonQuery(@"
CREATE TABLE IF NOT EXISTS metadata (
    key TEXT PRIMARY KEY,
    value TEXT
);
CREATE TABLE IF NOT EXISTS folders (
    id INTEGER PRIMARY KEY,
    store_id TEXT NOT NULL,
    entry_id TEXT NOT NULL,
    folder_path TEXT NOT NULL,
    UNIQUE(store_id, entry_id)
);
CREATE TABLE IF NOT EXISTS sender_folder_stats (
    sender_address TEXT NOT NULL COLLATE NOCASE,
    folder_id INTEGER NOT NULL,
    message_count INTEGER NOT NULL,
    last_seen_utc TEXT NOT NULL,
    PRIMARY KEY(sender_address, folder_id),
    FOREIGN KEY(folder_id) REFERENCES folders(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS subject_word_folder_stats (
    word TEXT NOT NULL COLLATE NOCASE,
    folder_id INTEGER NOT NULL,
    message_count INTEGER NOT NULL,
    last_seen_utc TEXT NOT NULL,
    PRIMARY KEY(word, folder_id),
    FOREIGN KEY(folder_id) REFERENCES folders(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_stats_sender_count
ON sender_folder_stats(sender_address, message_count DESC);
CREATE INDEX IF NOT EXISTS idx_stats_folder
ON sender_folder_stats(folder_id);
CREATE INDEX IF NOT EXISTS idx_word_stats_word_count
ON subject_word_folder_stats(word, message_count DESC);
CREATE INDEX IF NOT EXISTS idx_word_stats_folder
ON subject_word_folder_stats(folder_id);" );
            using (SQLiteTransaction schema = connection.BeginTransaction())
            {
                SetMetadata("schema_version", "2", schema);
                schema.Commit();
            }
        }

        private void EnsureIncrementCommand()
        {
            if (incrementCommand != null)
            {
                return;
            }

            incrementCommand = connection.CreateCommand();
            incrementCommand.Transaction = transaction;
            incrementCommand.CommandText = @"
INSERT INTO sender_folder_stats(sender_address, folder_id, message_count, last_seen_utc)
VALUES(@sender, @folder_id, 1, @seen_utc)
ON CONFLICT(sender_address, folder_id) DO UPDATE SET
    message_count = message_count + 1,
    last_seen_utc = excluded.last_seen_utc;";
            incrementCommand.Parameters.Add("@sender", System.Data.DbType.String);
            incrementCommand.Parameters.Add("@folder_id", System.Data.DbType.Int64);
            incrementCommand.Parameters.Add("@seen_utc", System.Data.DbType.String);
        }

        private void EnsureIncrementWordCommand()
        {
            if (incrementWordCommand != null)
            {
                return;
            }

            incrementWordCommand = connection.CreateCommand();
            incrementWordCommand.Transaction = transaction;
            incrementWordCommand.CommandText = @"
INSERT INTO subject_word_folder_stats(word, folder_id, message_count, last_seen_utc)
VALUES(@word, @folder_id, 1, @seen_utc)
ON CONFLICT(word, folder_id) DO UPDATE SET
    message_count = message_count + 1,
    last_seen_utc = excluded.last_seen_utc;";
            incrementWordCommand.Parameters.Add("@word", System.Data.DbType.String);
            incrementWordCommand.Parameters.Add("@folder_id", System.Data.DbType.Int64);
            incrementWordCommand.Parameters.Add("@seen_utc", System.Data.DbType.String);
        }

        private void RegisterChange()
        {
            changesSinceCommit++;
            if (changesSinceCommit >= CommitInterval)
            {
                CommitAndContinue();
            }
        }

        private void BeginTransaction()
        {
            transaction = connection.BeginTransaction();
            changesSinceCommit = 0;
        }

        private void CommitAndContinue()
        {
            incrementCommand?.Dispose();
            incrementCommand = null;
            incrementWordCommand?.Dispose();
            incrementWordCommand = null;
            transaction.Commit();
            transaction.Dispose();
            transaction = null;
            BeginTransaction();
        }

        private void CommitCurrentTransaction()
        {
            incrementCommand?.Dispose();
            incrementCommand = null;
            incrementWordCommand?.Dispose();
            incrementWordCommand = null;
            if (transaction != null)
            {
                transaction.Commit();
                transaction.Dispose();
                transaction = null;
            }
        }

        private void SetMetadata(string key, string value, SQLiteTransaction activeTransaction)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = activeTransaction;
                command.CommandText = @"
INSERT INTO metadata(key, value) VALUES(@key, @value)
ON CONFLICT(key) DO UPDATE SET value = excluded.value;";
                command.Parameters.AddWithValue("@key", key);
                command.Parameters.AddWithValue("@value", (object)value ?? DBNull.Value);
                command.ExecuteNonQuery();
            }
        }

        private void ExecuteNonQuery(string sql, SQLiteTransaction activeTransaction = null)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = activeTransaction;
                command.CommandText = sql;
                command.ExecuteNonQuery();
            }
        }

        private sealed class SenderExport
        {
            public SenderExport(string address)
            {
                Address = address;
                Folders = new System.Collections.Generic.List<FolderExport>();
            }

            public string Address { get; }
            public System.Collections.Generic.List<FolderExport> Folders { get; }
        }

        private sealed class FolderExport
        {
            public FolderExport(string path, long count)
            {
                Path = path;
                Count = count;
            }

            public string Path { get; }
            public long Count { get; }
        }

        private sealed class WordExport
        {
            public WordExport(string word)
            {
                Word = word;
                Folders = new System.Collections.Generic.List<FolderExport>();
            }

            public string Word { get; }
            public System.Collections.Generic.List<FolderExport> Folders { get; }
        }
    }
}
