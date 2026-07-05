using System;
using System.IO;
using FluentAssertions;
using RBLclass.Core;
using RBLclass.Core.Persistence;
using Xunit;

namespace RBLclass.Core.Tests
{
    public sealed class SqliteSettingsStoreTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly SqliteSettingsStore _store;

        public SqliteSettingsStoreTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(),
                "rblclass-set-" + Guid.NewGuid().ToString("N") + ".db");
            _store = new SqliteSettingsStore("Data Source=" + _dbPath);
            _store.EnsureSchema();
        }

        public void Dispose()
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
            foreach (var p in new[] { _dbPath, _dbPath + "-wal", _dbPath + "-shm" })
            {
                try { if (File.Exists(p)) File.Delete(p); } catch { }
            }
        }

        [Fact]
        public void EnsureSchema_is_idempotent()
        {
            _store.EnsureSchema();
            _store.Get("missing", "fallback").Should().Be("fallback");
        }

        [Fact]
        public void Get_returns_fallback_when_absent()
        {
            _store.Get("nope").Should().BeNull();
            _store.Get("nope", "x").Should().Be("x");
            _store.GetBool("nope", true).Should().BeTrue();
            _store.GetBool("nope", false).Should().BeFalse();
        }

        [Fact]
        public void Set_then_Get_round_trips_and_overwrites()
        {
            _store.Set("k", "one");
            _store.Get("k").Should().Be("one");

            _store.Set("k", "two");
            _store.Get("k").Should().Be("two");
        }

        [Fact]
        public void Bool_round_trips_as_one_zero()
        {
            _store.SetBool(SettingsKeys.OpenInNewWindow, true);
            _store.Get(SettingsKeys.OpenInNewWindow).Should().Be("1");
            _store.GetBool(SettingsKeys.OpenInNewWindow, false).Should().BeTrue();

            _store.SetBool(SettingsKeys.OpenInNewWindow, false);
            _store.GetBool(SettingsKeys.OpenInNewWindow, true).Should().BeFalse();
        }

        [Fact]
        public void Value_survives_a_new_store_instance_same_db()
        {
            _store.SetBool(SettingsKeys.AllResults, true);

            var reopened = new SqliteSettingsStore("Data Source=" + _dbPath);
            reopened.GetBool(SettingsKeys.AllResults, false).Should().BeTrue();
        }

        [Fact]
        public void Open_sets_WAL_journal_mode()
        {
            // EnsureSchema (ctor) already opened a connection; WAL is persisted
            // in the file header. Query via a raw connection to prove it stuck.
            using (var conn = new Microsoft.Data.Sqlite.SqliteConnection("Data Source=" + _dbPath))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA journal_mode;";
                    ((string)cmd.ExecuteScalar()).Should().Be("wal");
                }
            }
        }

        [Fact]
        public void GetAll_returns_empty_on_a_fresh_schema()
        {
            _store.GetAll().Should().BeEmpty();
        }

        [Fact]
        public void GetAll_returns_every_stored_entry()
        {
            _store.Set("a", "alpha");
            _store.Set("b", "bravo");
            _store.SetBool(SettingsKeys.KeepCopy, true);

            var all = _store.GetAll();

            all.Should().ContainKey("a").WhoseValue.Should().Be("alpha");
            all.Should().ContainKey("b").WhoseValue.Should().Be("bravo");
            all.Should().ContainKey(SettingsKeys.KeepCopy).WhoseValue.Should().Be("1");
        }

        [Fact]
        public void SetMany_writes_all_entries_and_round_trips_via_Get_and_GetAll()
        {
            var pairs = new[]
            {
                new System.Collections.Generic.KeyValuePair<string, string>("x", "10"),
                new System.Collections.Generic.KeyValuePair<string, string>("y", "20"),
                new System.Collections.Generic.KeyValuePair<string, string>("z", null),
            };

            _store.SetMany(pairs);

            _store.Get("x").Should().Be("10");
            _store.Get("y").Should().Be("20");
            _store.Get("z").Should().BeNull();

            var all = _store.GetAll();
            all.Should().ContainKey("x").WhoseValue.Should().Be("10");
            all.Should().ContainKey("y").WhoseValue.Should().Be("20");
            all.Should().ContainKey("z").WhoseValue.Should().BeNull();
        }

        [Fact]
        public void SetMany_overwrites_existing_keys()
        {
            _store.Set("k", "old");
            _store.SetMany(new[]
            {
                new System.Collections.Generic.KeyValuePair<string, string>("k", "new"),
            });
            _store.Get("k").Should().Be("new");
        }
    }
}
