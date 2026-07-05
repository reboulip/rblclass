using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;

namespace RBLclass.Core.Persistence
{
    /// <summary>
    /// SQLite-backed <see cref="ISettingsStore"/> over the <c>Settings</c>
    /// (Key, Value) table. Shares the database with the folder index; one
    /// connection per operation.
    /// </summary>
    public sealed class SqliteSettingsStore : ISettingsStore
    {
        /// <summary>
        /// The Settings table DDL, shared with <see cref="SqliteFolderRepository"/>'s
        /// v1 migration so the two cannot drift (v2.7.1.0 E2).
        /// </summary>
        internal const string CreateTableSql =
            "CREATE TABLE IF NOT EXISTS Settings (" +
            "  Key   TEXT PRIMARY KEY," +
            "  Value TEXT" +
            ");";

        private readonly string _connectionString;

        public SqliteSettingsStore(string connectionString)
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

        public string Get(string key, string fallback = null)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));

            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Value FROM Settings WHERE Key = $k;";
                cmd.Parameters.AddWithValue("$k", key);
                var value = cmd.ExecuteScalar();
                return value == null || value is DBNull ? fallback : (string)value;
            }
        }

        public void Set(string key, string value)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));

            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "INSERT OR REPLACE INTO Settings (Key, Value) VALUES ($k, $v);";
                cmd.Parameters.AddWithValue("$k", key);
                cmd.Parameters.AddWithValue("$v", (object)value ?? DBNull.Value);
                cmd.ExecuteNonQuery();
            }
        }

        public bool GetBool(string key, bool fallback)
        {
            var raw = Get(key, null);
            if (raw == null) return fallback;
            if (raw == "1") return true;
            if (raw == "0") return false;
            return bool.TryParse(raw, out var parsed) ? parsed : fallback;
        }

        public void SetBool(string key, bool value) => Set(key, value ? "1" : "0");

        public IDictionary<string, string> GetAll()
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            using (var conn = Open())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Key, Value FROM Settings;";
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                        result[reader.GetString(0)] = reader.IsDBNull(1) ? null : reader.GetString(1);
                }
            }
            return result;
        }

        public void SetMany(IEnumerable<KeyValuePair<string, string>> entries)
        {
            if (entries == null) throw new ArgumentNullException(nameof(entries));

            using (var conn = Open())
            using (var tx = conn.BeginTransaction())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText =
                        "INSERT OR REPLACE INTO Settings (Key, Value) VALUES ($k, $v);";
                    var pKey = cmd.Parameters.Add("$k", SqliteType.Text);
                    var pVal = cmd.Parameters.Add("$v", SqliteType.Text);

                    foreach (var entry in entries)
                    {
                        if (entry.Key == null) continue;
                        pKey.Value = entry.Key;
                        pVal.Value = (object)entry.Value ?? DBNull.Value;
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

            // WAL for concurrent reads during a write, matching the other two
            // stores on this shared DB (v2.7.1.0 E2).
            using (var pragma = conn.CreateCommand())
            {
                pragma.CommandText = "PRAGMA journal_mode=WAL;";
                pragma.ExecuteNonQuery();
            }

            return conn;
        }
    }
}
