using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Presents a favourite-folder snapshot as an <see cref="IFolderTree"/> so the
    /// existing <see cref="FolderSearchService"/> matcher (AND keywords,
    /// case/accent-insensitive) searches favourites exactly like Outlook folders
    /// (v2.4.0.0 F1). Each favourite becomes a synthetic leaf <see cref="FolderNode"/>
    /// keyed by its path; only <see cref="GetAll"/> is used by the searcher.
    /// </summary>
    internal sealed class FavoriteFolderSearchAdapter : IFolderTree
    {
        public const string FavoritesStoreId = "__favorites__";

        private readonly IReadOnlyList<FolderNode> _nodes;

        public FavoriteFolderSearchAdapter(IReadOnlyList<FavoriteFolder> favorites)
        {
            _nodes = favorites.Select(f => new FolderNode(
                storeId: FavoritesStoreId,
                entryId: f.Path,
                parentEntryId: null, // flat list: every directory is independently selectable
                name: f.DisplayName,
                fullPath: f.Path,
                isLeaf: true)).ToArray();
        }

        public IReadOnlyList<FolderNode> GetAll() => _nodes;

        // The index-lifecycle methods are never called on this read-only adapter.
        public IndexResult Load() => throw new NotSupportedException();
        public IndexResult WalkAndPersist() => throw new NotSupportedException();
        public void ReindexStore(string storeId) => throw new NotSupportedException();
    }
}
