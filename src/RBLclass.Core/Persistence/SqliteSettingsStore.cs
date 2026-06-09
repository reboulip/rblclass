using System;
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
                cmd.CommandText =
                    "CREATE TABLE IF NOT EXISTS Settings (" +
                    "  Key   TEXT PRIMARY KEY," +
                    "  Value TEXT" +
                    ");";
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

        private SqliteConnection Open()
        {
            var conn = new SqliteConnection(_connectionString);
            conn.Open();
            return conn;
        }
    }
}
