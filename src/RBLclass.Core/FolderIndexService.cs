using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Orchestrates the folder index lifecycle over an <see cref="IMailStore"/>
    /// (live Outlook) and an <see cref="IFolderRepository"/> (SQLite cache),
    /// applying the store-level <see cref="FolderExclusionPolicy"/>. This is the
    /// reimplementation of the legacy <c>indexFolders</c>, but persistent: the
    /// tree is walked once and thereafter loaded from SQLite.
    /// </summary>
    public sealed class FolderIndexService : IFolderIndexService
    {
        private readonly IMailStore _mailStore;
        private readonly IFolderRepository _repository;
        private readonly FolderExclusionPolicy _exclusion;

        private readonly object _gate = new object();
        private List<FolderNode> _cache = new List<FolderNode>();
        private IndexStatus _indexStatus = IndexStatus.NotFound;

        public event PropertyChangedEventHandler PropertyChanged;

        public IndexStatus IndexStatus
        {
            get { lock (_gate) { return _indexStatus; } }
            private set
            {
                lock (_gate) { _indexStatus = value; }
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IndexStatus)));
            }
        }

        public FolderIndexService(IMailStore mailStore,
                                  IFolderRepository repository,
                                  FolderExclusionOptions exclusionOptions = null)
        {
            _mailStore = mailStore ?? throw new ArgumentNullException(nameof(mailStore));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _exclusion = new FolderExclusionPolicy(
                exclusionOptions ?? FolderExclusionOptions.Default);
        }

        public IndexResult Load()
        {
            _repository.EnsureSchema();

            if (!_repository.HasAnyFolders())
            {
                IndexStatus = IndexStatus.NotFound;
                return new IndexResult(IndexSource.NeedsWalk, 0, 0);
            }

            var folders = _repository.LoadAll().ToList();
            SetCache(folders);
            IndexStatus = IndexStatus.Ready;
            return new IndexResult(IndexSource.LoadedFromCache,
                                   CountStores(folders), folders.Count);
        }

        public IndexResult WalkAndPersist()
        {
            IndexStatus = IndexStatus.Indexing;
            _repository.EnsureSchema();

            var all = new List<FolderNode>();
            int storeCount = 0;

            foreach (var store in _mailStore.GetStores())
            {
                if (_exclusion.IsStoreExcluded(store))
                    continue;

                storeCount++;
                all.AddRange(_mailStore.GetFolders(store.StoreId));
            }

            _repository.ReplaceAll(all);
            SetCache(all);
            IndexStatus = IndexStatus.Ready;
            return new IndexResult(IndexSource.Walked, storeCount, all.Count);
        }

        public void MarkIndexing() => IndexStatus = IndexStatus.Indexing;

        public void MarkReady() => IndexStatus = IndexStatus.Ready;

        public IndexResult PersistWalkedStores(
            IReadOnlyList<(StoreInfo Store, IReadOnlyList<FolderNode> Folders)> walkedStores)
        {
            if (walkedStores == null) throw new ArgumentNullException(nameof(walkedStores));
            _repository.EnsureSchema();
            var all = walkedStores.SelectMany(w => w.Folders).ToList();
            _repository.ReplaceAll(all);
            SetCache(all);
            return new IndexResult(IndexSource.Walked, walkedStores.Count, all.Count);
        }

        public void ReindexStore(string storeId)
        {
            if (storeId == null) throw new ArgumentNullException(nameof(storeId));

            _repository.EnsureSchema();

            var folders = _mailStore.GetFolders(storeId).ToList();
            _repository.ReplaceStore(storeId, folders);

            lock (_gate)
            {
                var merged = _cache.Where(f => f.StoreId != storeId).ToList();
                merged.AddRange(folders);
                _cache = merged;
            }
        }

        public IReadOnlyList<FolderNode> GetAll()
        {
            lock (_gate)
            {
                return _cache.ToArray();
            }
        }

        private void SetCache(List<FolderNode> folders)
        {
            lock (_gate)
            {
                _cache = folders;
            }
        }

        private static int CountStores(IEnumerable<FolderNode> folders)
        {
            return folders.Select(f => f.StoreId).Distinct().Count();
        }
    }
}
