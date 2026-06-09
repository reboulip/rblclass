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
    }
}
