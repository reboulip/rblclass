using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Data.Sqlite;

namespace RBLclass.Core.Persistence
{
    /// <summary>
    /// SQLite-backed <see cref="IClassificationHistory"/> over the
    /// <c>ClassificationHistory</c> table (schema v2). Shares the database with
    /// the folder index and settings; one connection per operation, WAL, per
    /// CLAUDE.md SQLite rules. <see cref="EnsureSchema"/> is idempotent and
    /// independent of the folder repository's migration (same DDL).
    /// </summary>
    public sealed class SqliteClassificationHistory : IClassificationHistory
    {
        private readonly string _connectionString;

        public SqliteClassificationHistory(string connectionString)
        {
            _connectionString = connectionString
                ?? throw new ArgumentNullException(nameof(connectionString));
        }

        public void EnsureSchema()
        {
            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = CreateTableSql;
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>Shared with <see cref="SqliteFolderRepository"/>'s v2 migration - keep both in sync.</summary>
        internal const string CreateTableSql =
            "CREATE TABLE IF NOT EXISTS ClassificationHistory (" +
            "  Id              INTEGER PRIMARY KEY AUTOINCREMENT," +
            "  BatchId         TEXT NOT NULL," +
            "  ConversationKey TEXT NOT NULL," +
            "  DestStoreId     TEXT NOT NULL," +
            "  DestEntryId     TEXT NOT NULL," +
            "  WhenUtc         TEXT NOT NULL" +
            ");" +
            "CREATE INDEX IF NOT EXISTS IX_ClassificationHistory_Key " +
            "ON ClassificationHistory(ConversationKey);";

        public void Record(string batchId, string conversationKey,
                           IEnumerable<FolderNode> destinations, DateTime whenUtc)
        {
            if (batchId == null) throw new ArgumentNullException(nameof(batchId));
            if (conversationKey == null) throw new ArgumentNullException(nameof(conversationKey));
            if (destinations == null) throw new ArgumentNullException(nameof(destinations));

            using (var conn = Open())
            using (var tx = conn.BeginTransaction())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText =
                        "INSERT INTO ClassificationHistory " +
                        "(BatchId, ConversationKey, DestStoreId, DestEntryId, WhenUtc) " +
                        "VALUES ($batch, $key, $storeId, $entryId, $when);";

                    var pBatch = cmd.Parameters.Add("$batch", SqliteType.Text);
                    var pKey = cmd.Parameters.Add("$key", SqliteType.Text);
                    var pStore = cmd.Parameters.Add("$storeId", SqliteType.Text);
                    var pEntry = cmd.Parameters.Add("$entryId", SqliteType.Text);
                    var pWhen = cmd.Parameters.Add("$when", SqliteType.Text);

                    pBatch.Value = batchId;
                    pKey.Value = conversationKey;
                    pWhen.Value = whenUtc.ToString("o", CultureInfo.InvariantCulture);

                    foreach (var destination in destinations)
                    {
                        pStore.Value = destination.StoreId;
                        pEntry.Value = destination.EntryId;
                        cmd.ExecuteNonQuery();
                    }
                }
                tx.Commit();
            }
        }

        public IReadOnlyList<HistoryDestination> GetLatestDestinations(string conversationKey)
        {
            var result = new List<HistoryDestination>();
            if (string.IsNullOrEmpty(conversationKey)) return result;

            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                // The latest classify action for this conversation is the batch
                // of its highest row id; return that whole destination set.
                cmd.CommandText =
                    "SELECT DestStoreId, DestEntryId FROM ClassificationHistory " +
                    "WHERE ConversationKey = $key AND BatchId = (" +
                    "  SELECT BatchId FROM ClassificationHistory " +
                    "  WHERE ConversationKey = $key ORDER BY Id DESC LIMIT 1" +
                    ") ORDER BY Id;";
                cmd.Parameters.AddWithValue("$key", conversationKey);

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                        result.Add(new HistoryDestination(reader.GetString(0), reader.GetString(1)));
                }
            }

            return result;
        }

        public void DeleteBatches(IEnumerable<string> batchIds)
        {
            if (batchIds == null) return;

            using (var conn = Open())
            using (var tx = conn.BeginTransaction())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "DELETE FROM ClassificationHistory WHERE BatchId = $batch;";
                    var pBatch = cmd.Parameters.Add("$batch", SqliteType.Text);

                    foreach (var batchId in batchIds)
                    {
                        if (batchId == null) continue;
                        pBatch.Value = batchId;
                        cmd.ExecuteNonQuery();
                    }
                }
                tx.Commit();
            }
        }

        private SqliteConnection Open()
        {
            var conn = new SqliteConnection(_connectionString);
            conn.Open();

            using (var pragma = conn.CreateCommand())
            {
                pragma.CommandText = "PRAGMA journal_mode=WAL;";
                pragma.ExecuteNonQuery();
            }

            return conn;
        }
    }
}
