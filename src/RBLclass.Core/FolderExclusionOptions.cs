using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Configurable rules for which stores/folders to skip during indexing.
    /// Replaces the legacy hard-coded FR/EN substrings ("Dossiers publics",
    /// "G-FRA", "Éléments supprimés") with data (deviation #8). Immutable;
    /// build a new instance to change the policy.
    /// </summary>
    public sealed class FolderExclusionOptions
    {
        public FolderExclusionOptions(
            bool excludePublicFolderStores = true,
            bool excludeDeletedItems = true,
            IEnumerable<string> excludedStoreNameSubstrings = null)
        {
            ExcludePublicFolderStores = excludePublicFolderStores;
            ExcludeDeletedItems = excludeDeletedItems;
            ExcludedStoreNameSubstrings =
                (excludedStoreNameSubstrings ?? Enumerable.Empty<string>())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToArray();
        }

        /// <summary>
        /// When true (default), Exchange public-folder stores are skipped. Every
        /// other store - PST archives AND the primary mailbox (Exchange/IMAP/OST)
        /// - is indexed, matching the legacy tool, which walked all open stores
        /// and only excluded public folders.
        /// </summary>
        public bool ExcludePublicFolderStores { get; }

        /// <summary>
        /// When true (default), the Deleted Items subtree of every store is
        /// pruned from the index.
        /// </summary>
        public bool ExcludeDeletedItems { get; }

        /// <summary>
        /// Additional case-insensitive store-name substrings to exclude. Empty
        /// by default; this is where a user re-adds locale-specific names if
        /// <see cref="OnlyDataFileStores"/> is not enough.
        /// </summary>
        public IReadOnlyList<string> ExcludedStoreNameSubstrings { get; }

        /// <summary>
        /// The default policy: index all stores except Exchange public folders,
        /// with each store's Deleted Items subtree pruned.
        /// </summary>
        public static FolderExclusionOptions Default { get; } = new FolderExclusionOptions();
    }
}
