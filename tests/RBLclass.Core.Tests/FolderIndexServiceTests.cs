using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using RBLclass.Core.Persistence;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// The index-lifecycle surface: first run walks Outlook and persists;
    /// every later start loads from SQLite without touching Outlook. Uses a
    /// faked <see cref="IMailStore"/> (NSubstitute) and a real temp-file
    /// repository so the walk-vs-load decision is exercised against actual
    /// persistence.
    /// </summary>
    public sealed class FolderIndexServiceTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public FolderIndexServiceTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(),
                "rblclass-svc-" + Guid.NewGuid().ToString("N") + ".db");
            _connectionString = "Data Source=" + _dbPath;
        }

        public void Dispose()
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
            foreach (var path in new[] { _dbPath, _dbPath + "-wal", _dbPath + "-shm" })
            {
                try { if (File.Exists(path)) File.Delete(path); }
                catch { /* best effort */ }
            }
        }

        private IFolderRepository NewRepo() => new SqliteFolderRepository(_connectionString);

        private static FolderNode Node(string store, string entry, string parent,
                                       string name, string path, bool leaf = true)
            => new FolderNode(store, entry, parent, name, path, leaf);

        private static IMailStore StoreWith(
            IEnumerable<StoreInfo> stores,
            Dictionary<string, IReadOnlyList<FolderNode>> foldersByStore)
        {
            var mock = Substitute.For<IMailStore>();
            mock.GetStores().Returns(stores.ToArray());
            foreach (var kv in foldersByStore)
                mock.GetFolders(kv.Key).Returns(kv.Value);
            return mock;
        }

        [Fact]
        public void Load_on_empty_index_reports_NeedsWalk_and_does_not_touch_outlook()
        {
            var mail = Substitute.For<IMailStore>();
            var service = new FolderIndexService(mail, NewRepo());

            var result = service.Load();

            result.Source.Should().Be(IndexSource.NeedsWalk);
            service.GetAll().Should().BeEmpty();
            mail.DidNotReceive().GetStores();
        }

        [Fact]
        public void WalkAndPersist_walks_included_stores_and_skips_excluded()
        {
            var stores = new[]
            {
                new StoreInfo("pst1", "Archive 2024", isDataFileStore: true),
                new StoreInfo("pub", "Public Folders", isDataFileStore: false,
                              isPublicFolderStore: true), // excluded by default
            };
            var folders = new Dictionary<string, IReadOnlyList<FolderNode>>
            {
                ["pst1"] = new[]
                {
                    Node("pst1", "e1", null, "Inbox", "Archive 2024 / Inbox", leaf: false),
                    Node("pst1", "e2", "e1", "Projects", "Archive 2024 / Inbox / Projects"),
                },
            };
            var mail = StoreWith(stores, folders);
            var service = new FolderIndexService(mail, NewRepo());

            var result = service.WalkAndPersist();

            result.Source.Should().Be(IndexSource.Walked);
            result.StoreCount.Should().Be(1);          // public-folder store skipped
            result.FolderCount.Should().Be(2);
            service.GetAll().Select(f => f.EntryId).Should().BeEquivalentTo("e1", "e2");
            mail.DidNotReceive().GetFolders("pub");    // never walked the excluded store
        }

        [Fact]
        public void Second_session_loads_from_cache_without_walking_outlook()
        {
            var stores = new[] { new StoreInfo("pst1", "Archive", isDataFileStore: true) };
            var folders = new Dictionary<string, IReadOnlyList<FolderNode>>
            {
                ["pst1"] = new[] { Node("pst1", "e1", null, "Inbox", "Archive / Inbox") },
            };

            // First session: walk + persist.
            new FolderIndexService(StoreWith(stores, folders), NewRepo()).WalkAndPersist();

            // Second session: a brand-new service + a mail store that would
            // throw if walked, proving the load path never touches Outlook.
            var coldMail = Substitute.For<IMailStore>();
            var service2 = new FolderIndexService(coldMail, NewRepo());

            var result = service2.Load();

            result.Source.Should().Be(IndexSource.LoadedFromCache);
            result.StoreCount.Should().Be(1);
            result.FolderCount.Should().Be(1);
            service2.GetAll().Should().ContainSingle().Which.EntryId.Should().Be("e1");
            coldMail.DidNotReceive().GetStores();
            coldMail.DidNotReceive().GetFolders(Arg.Any<string>());
        }

        [Fact]
        public void ReindexStore_replaces_only_that_store_in_cache_and_db()
        {
            var stores = new[]
            {
                new StoreInfo("s1", "One", isDataFileStore: true),
                new StoreInfo("s2", "Two", isDataFileStore: true),
            };
            var folders = new Dictionary<string, IReadOnlyList<FolderNode>>
            {
                ["s1"] = new[] { Node("s1", "a1", null, "A1", "One / A1") },
                ["s2"] = new[] { Node("s2", "b1", null, "B1", "Two / B1") },
            };
            var mail = StoreWith(stores, folders);
            var service = new FolderIndexService(mail, NewRepo());
            service.WalkAndPersist();

            // s1 gains a sub-folder; re-walk just s1.
            mail.GetFolders("s1").Returns(new[]
            {
                Node("s1", "a1", null, "A1", "One / A1", leaf: false),
                Node("s1", "a2", "a1", "A2", "One / A1 / A2"),
            });

            service.ReindexStore("s1");

            service.GetAll().Where(f => f.StoreId == "s1").Select(f => f.EntryId)
                   .Should().BeEquivalentTo("a1", "a2");
            service.GetAll().Where(f => f.StoreId == "s2").Select(f => f.EntryId)
                   .Should().BeEquivalentTo("b1");

            // And it persisted: a fresh service loads the same shape.
            var reloaded = new FolderIndexService(Substitute.For<IMailStore>(), NewRepo());
            reloaded.Load();
            reloaded.GetAll().Select(f => f.EntryId)
                    .Should().BeEquivalentTo("a1", "a2", "b1");
        }

        [Fact]
        public void IndexStatus_starts_NotFound_and_stays_NotFound_after_empty_Load()
        {
            var service = new FolderIndexService(Substitute.For<IMailStore>(), NewRepo());

            service.IndexStatus.Should().Be(IndexStatus.NotFound);

            service.Load();

            service.IndexStatus.Should().Be(IndexStatus.NotFound);
        }

        [Fact]
        public void Load_from_cache_sets_IndexStatus_Ready()
        {
            var stores = new[] { new StoreInfo("pst1", "Archive", isDataFileStore: true) };
            var folders = new Dictionary<string, IReadOnlyList<FolderNode>>
            {
                ["pst1"] = new[] { Node("pst1", "e1", null, "Inbox", "Archive / Inbox") },
            };
            new FolderIndexService(StoreWith(stores, folders), NewRepo()).WalkAndPersist();

            var service = new FolderIndexService(Substitute.For<IMailStore>(), NewRepo());
            service.Load();

            service.IndexStatus.Should().Be(IndexStatus.Ready);
        }

        [Fact]
        public void WalkAndPersist_transitions_Indexing_then_Ready_and_raises_PropertyChanged()
        {
            var stores = new[] { new StoreInfo("pst1", "Archive", isDataFileStore: true) };
            var folders = new Dictionary<string, IReadOnlyList<FolderNode>>
            {
                ["pst1"] = new[] { Node("pst1", "e1", null, "Inbox", "Archive / Inbox") },
            };
            var service = new FolderIndexService(StoreWith(stores, folders), NewRepo());

            var observed = new List<IndexStatus>();
            service.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(FolderIndexService.IndexStatus))
                    observed.Add(service.IndexStatus);
            };

            service.WalkAndPersist();

            observed.Should().ContainInOrder(IndexStatus.Indexing, IndexStatus.Ready);
            service.IndexStatus.Should().Be(IndexStatus.Ready);
        }
    }
}
