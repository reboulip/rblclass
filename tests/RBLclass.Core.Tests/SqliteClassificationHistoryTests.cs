using System;
using System.IO;
using FluentAssertions;
using RBLclass.Core;
using RBLclass.Core.Persistence;
using Xunit;

namespace RBLclass.Core.Tests
{
    public sealed class SqliteClassificationHistoryTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly SqliteClassificationHistory _history;

        public SqliteClassificationHistoryTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(),
                "rblclass-history-" + Guid.NewGuid().ToString("N") + ".db");
            _history = new SqliteClassificationHistory("Data Source=" + _dbPath);
            _history.EnsureSchema();
        }

        public void Dispose()
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
            foreach (var p in new[] { _dbPath, _dbPath + "-wal", _dbPath + "-shm" })
            {
                try { if (File.Exists(p)) File.Delete(p); } catch { }
            }
        }

        private static FolderNode Dest(string store, string entry) =>
            new FolderNode(store, entry, null, entry, entry, isLeaf: true);

        [Fact]
        public void Unknown_conversation_has_no_destinations()
        {
            _history.GetLatestDestinations("never-seen").Should().BeEmpty();
        }

        [Fact]
        public void Records_and_returns_the_destination_set_of_a_conversation()
        {
            _history.Record("b1", "conv-1",
                new[] { Dest("s1", "d1"), Dest("s1", "d2") }, DateTime.UtcNow);

            var latest = _history.GetLatestDestinations("conv-1");

            latest.Should().HaveCount(2);
            latest.Should().Contain(d => d.StoreId == "s1" && d.EntryId == "d1");
            latest.Should().Contain(d => d.StoreId == "s1" && d.EntryId == "d2");
        }

        [Fact]
        public void Latest_batch_wins_for_a_conversation()
        {
            _history.Record("b1", "conv-1", new[] { Dest("s1", "old") }, DateTime.UtcNow);
            _history.Record("b2", "conv-1", new[] { Dest("s1", "new1"), Dest("s1", "new2") },
                DateTime.UtcNow);

            var latest = _history.GetLatestDestinations("conv-1");

            // Only the most recent classify's destinations, not a union.
            latest.Should().HaveCount(2);
            latest.Should().OnlyContain(d => d.EntryId == "new1" || d.EntryId == "new2");
        }

        [Fact]
        public void Conversations_are_independent()
        {
            _history.Record("b1", "conv-1", new[] { Dest("s1", "a") }, DateTime.UtcNow);
            _history.Record("b2", "conv-2", new[] { Dest("s1", "b") }, DateTime.UtcNow);

            _history.GetLatestDestinations("conv-1").Should().ContainSingle()
                    .Which.EntryId.Should().Be("a");
            _history.GetLatestDestinations("conv-2").Should().ContainSingle()
                    .Which.EntryId.Should().Be("b");
        }

        [Fact]
        public void Deleting_a_batch_rolls_back_to_the_previous_one()
        {
            _history.Record("b1", "conv-1", new[] { Dest("s1", "old") }, DateTime.UtcNow);
            _history.Record("b2", "conv-1", new[] { Dest("s1", "new") }, DateTime.UtcNow);

            _history.DeleteBatches(new[] { "b2" });

            // With the latest batch gone, the earlier one surfaces again.
            _history.GetLatestDestinations("conv-1").Should().ContainSingle()
                    .Which.EntryId.Should().Be("old");
        }

        [Fact]
        public void Deleting_the_only_batch_leaves_no_history()
        {
            _history.Record("b1", "conv-1", new[] { Dest("s1", "a") }, DateTime.UtcNow);

            _history.DeleteBatches(new[] { "b1" });

            _history.GetLatestDestinations("conv-1").Should().BeEmpty();
        }
    }
}
