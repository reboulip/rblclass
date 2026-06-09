using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using RBLclass.Core;
using RBLclass.Core.Persistence;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// Exercises the real SQLite repository over a temporary database file
    /// (the round-trip surface the reimplementation roadmap calls for). Each
    /// test gets its own file and deletes it on dispose.
    /// </summary>
    public sealed class SqliteFolderRepositoryTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly SqliteFolderRepository _repo;

        public SqliteFolderRepositoryTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(),
                "rblclass-test-" + Guid.NewGuid().ToString("N") + ".db");
            _repo = new SqliteFolderRepository("Data Source=" + _dbPath);
        }

        public void Dispose()
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
            foreach (var path in new[] { _dbPath, _dbPath + "-wal", _dbPath + "-shm" })
            {
                try { if (File.Exists(path)) File.Delete(path); }
                catch { /* best effort cleanup */ }
            }
        }

        private static FolderNode Node(string store, string entry, string parent,
                                       string name, string path, bool leaf = true)
            => new FolderNode(store, entry, parent, name, path, leaf);

        [Fact]
        public void EnsureSchema_is_idempotent()
        {
            _repo.EnsureSchema();
            _repo.EnsureSchema(); // second call must not throw or duplicate

            _repo.HasAnyFolders().Should().BeFalse();
        }

        [Fact]
        public void HasAnyFolders_is_false_on_empty_then_true_after_insert()
        {
            _repo.EnsureSchema();
            _repo.HasAnyFolders().Should().BeFalse();

            _repo.ReplaceAll(new[] { Node("s1", "e1", null, "Inbox", "Arch / Inbox") });

            _repo.HasAnyFolders().Should().BeTrue();
        }

        [Fact]
        public void ReplaceAll_round_trips_nodes_ordered_by_full_path()
        {
            _repo.EnsureSchema();

            var nodes = new[]
            {
                Node("s1", "e2", "e1", "Zebra", "Arch / Zebra"),
                Node("s1", "e1", null, "Inbox", "Arch / Inbox", leaf: false),
                Node("s1", "e3", "e1", "Apple", "Arch / Inbox / Apple"),
            };

            _repo.ReplaceAll(nodes);
            var loaded = _repo.LoadAll();

            loaded.Select(n => n.FullPath).Should().ContainInOrder(
                "Arch / Inbox", "Arch / Inbox / Apple", "Arch / Zebra");

            var inbox = loaded.Single(n => n.EntryId == "e1");
            inbox.ParentEntryId.Should().BeNull();
            inbox.IsLeaf.Should().BeFalse();
            inbox.Name.Should().Be("Inbox");
            inbox.StoreId.Should().Be("s1");

            loaded.Single(n => n.EntryId == "e2").ParentEntryId.Should().Be("e1");
        }

        [Fact]
        public void ReplaceAll_clears_previous_content()
        {
            _repo.EnsureSchema();
            _repo.ReplaceAll(new[] { Node("s1", "e1", null, "Old", "Arch / Old") });

            _repo.ReplaceAll(new[] { Node("s1", "e2", null, "New", "Arch / New") });

            var loaded = _repo.LoadAll();
            loaded.Should().ContainSingle().Which.EntryId.Should().Be("e2");
        }

        [Fact]
        public void ReplaceStore_replaces_only_the_named_store()
        {
            _repo.EnsureSchema();
            _repo.ReplaceAll(new[]
            {
                Node("s1", "a1", null, "A1", "A / 1"),
                Node("s2", "b1", null, "B1", "B / 1"),
            });

            _repo.ReplaceStore("s1", new[]
            {
                Node("s1", "a2", null, "A2", "A / 2"),
            });

            var loaded = _repo.LoadAll();
            loaded.Where(n => n.StoreId == "s1").Select(n => n.EntryId)
                  .Should().BeEquivalentTo("a2");
            loaded.Where(n => n.StoreId == "s2").Select(n => n.EntryId)
                  .Should().BeEquivalentTo("b1");
        }
    }
}
