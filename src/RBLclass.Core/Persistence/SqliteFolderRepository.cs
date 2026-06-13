using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;

namespace RBLclass.Core.Persistence
{
    /// <summary>
    /// SQLite-backed <see cref="IFolderRepository"/>. The only I/O code in the
    /// core, kept behind the interface so the business logic stays pure
    /// (see the Step 1 architecture decision). Opens one connection per logical
    /// operation with WAL journaling, per CLAUDE.md SQLite rules.
    /// </summary>
    public sealed class SqliteFolderRepository : IFolderRepository
    {
        /// <summary>Schema version this build expects. Bump with each migration.</summary>
        public const int CurrentSchemaVersion = 2;

        private readonly string _connectionString;

        /// <param name="connectionString">
        /// e.g. "Data Source=C:\Users\...\AppData\Local\RBLclass\rblclass.db".
        /// Tests pass a temp-file or shared in-memory source.
        /// </param>
        public SqliteFolderRepository(string connectionString)
        {
            _connectionString = connectionString
                ?? throw new ArgumentNullException(nameof(connectionString));
        }

        public void EnsureSchema()
        {
            using (var conn = Open())
            {
                int version = ReadSchemaVersion(conn);
                if (version >= CurrentSchemaVersion)
                    return;

                using (var tx = conn.BeginTransaction())
                {
                    if (version < 1)
                        ApplyV1(conn, tx);
                    if (version < 2)
                        ApplyV2(conn, tx);

                    WriteSchemaVersion(conn, tx, CurrentSchemaVersion);
                    tx.Commit();
                }
            }
        }

        public bool HasAnyFolders()
        {
            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT EXISTS(SELECT 1 FROM Folders);";
                return Convert.ToInt64(cmd.ExecuteScalar()) != 0;
            }
        }

        public IReadOnlyList<FolderNode> LoadAll()
        {
            var result = new List<FolderNode>();

            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT StoreId, EntryId, ParentEntryId, Name, FullPath, IsLeaf " +
                    "FROM Folders ORDER BY FullPath;";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new FolderNode(
                            storeId: reader.GetString(0),
                            entryId: reader.GetString(1),
                            parentEntryId: reader.IsDBNull(2) ? null : reader.GetString(2),
                            name: reader.GetString(3),
                            fullPath: reader.GetString(4),
                            isLeaf: reader.GetInt64(5) != 0));
                    }
                }
            }

            return result;
        }

        public void ReplaceAll(IEnumerable<FolderNode> folders)
        {
            if (folders == null) throw new ArgumentNullException(nameof(folders));

            using (var conn = Open())
            using (var tx = conn.BeginTransaction())
            {
                using (var del = conn.CreateCommand())
                {
                    del.Transaction = tx;
                    del.CommandText = "DELETE FROM Folders;";
                    del.ExecuteNonQuery();
                }

                InsertMany(conn, tx, folders);
                tx.Commit();
            }
        }

        public void ReplaceStore(string storeId, IEnumerable<FolderNode> folders)
        {
            if (storeId == null) throw new ArgumentNullException(nameof(storeId));
            if (folders == null) throw new ArgumentNullException(nameof(folders));

            using (var conn = Open())
            using (var tx = conn.BeginTransaction())
            {
                using (var del = conn.CreateCommand())
                {
                    del.Transaction = tx;
                    del.CommandText = "DELETE FROM Folders WHERE StoreId = $storeId;";
                    del.Parameters.AddWithValue("$storeId", storeId);
                    del.ExecuteNonQuery();
                }

                InsertMany(conn, tx, folders);
                tx.Commit();
            }
        }

        // --- internals ------------------------------------------------------

        private SqliteConnection Open()
        {
            var conn = new SqliteConnection(_connectionString);
            conn.Open();

            // WAL improves concurrent read/write; harmless (no-op) for the
            // in-memory source used by some tests.
            using (var pragma = conn.CreateCommand())
            {
                pragma.CommandText = "PRAGMA journal_mode=WAL;";
                pragma.ExecuteNonQuery();
            }

            return conn;
        }

        private static void InsertMany(SqliteConnection conn,
                                       SqliteTransaction tx,
                                       IEnumerable<FolderNode> folders)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText =
                    "INSERT OR REPLACE INTO Folders " +
                    "(StoreId, EntryId, ParentEntryId, Name, FullPath, IsLeaf) " +
                    "VALUES ($storeId, $entryId, $parentEntryId, $name, $fullPath, $isLeaf);";

                var pStore = cmd.Parameters.Add("$storeId", SqliteType.Text);
                var pEntry = cmd.Parameters.Add("$entryId", SqliteType.Text);
                var pParent = cmd.Parameters.Add("$parentEntryId", SqliteType.Text);
                var pName = cmd.Parameters.Add("$name", SqliteType.Text);
                var pPath = cmd.Parameters.Add("$fullPath", SqliteType.Text);
                var pLeaf = cmd.Parameters.Add("$isLeaf", SqliteType.Integer);

                foreach (var f in folders)
                {
                    pStore.Value = f.StoreId;
                    pEntry.Value = f.EntryId;
                    pParent.Value = (object)f.ParentEntryId ?? DBNull.Value;
                    pName.Value = f.Name;
                    pPath.Value = f.FullPath;
                    pLeaf.Value = f.IsLeaf ? 1 : 0;
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private static int ReadSchemaVersion(SqliteConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT name FROM sqlite_master " +
                    "WHERE type='table' AND name='SchemaVersion';";
                if (cmd.ExecuteScalar() == null)
                    return 0;
            }

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT COALESCE(MAX(Version), 0) FROM SchemaVersion;";
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private static void WriteSchemaVersion(SqliteConnection conn,
                                               SqliteTransaction tx,
                                               int version)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "INSERT INTO SchemaVersion (Version) VALUES ($v);";
                cmd.Parameters.AddWithValue("$v", version);
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>Schema v1: folder index + settings + the version table.</summary>
        private static void ApplyV1(SqliteConnection conn, SqliteTransaction tx)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText =
                    "CREATE TABLE IF NOT EXISTS SchemaVersion (" +
                    "  Version INTEGER NOT NULL" +
                    ");" +
                    "CREATE TABLE IF NOT EXISTS Folders (" +
                    "  StoreId       TEXT    NOT NULL," +
                    "  EntryId       TEXT    NOT NULL," +
                    "  ParentEntryId TEXT    NULL," +
                    "  Name          TEXT    NOT NULL," +
                    "  FullPath      TEXT    NOT NULL," +
                    "  IsLeaf        INTEGER NOT NULL," +
                    "  PRIMARY KEY (StoreId, EntryId)" +
                    ");" +
                    "CREATE INDEX IF NOT EXISTS IX_Folders_StoreId ON Folders(StoreId);" +
                    "CREATE TABLE IF NOT EXISTS Settings (" +
                    "  Key   TEXT PRIMARY KEY," +
                    "  Value TEXT" +
                    ");";
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>Schema v2: the classification history behind Auto-class/Undo (v2.2).</summary>
        private static void ApplyV2(SqliteConnection conn, SqliteTransaction tx)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = SqliteClassificationHistory.CreateTableSql;
                cmd.ExecuteNonQuery();
            }
        }
    }
}
