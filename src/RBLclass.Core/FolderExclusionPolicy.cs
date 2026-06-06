using System;
using System.Globalization;

namespace RBLclass.Core
{
    /// <summary>
    /// Pure decision logic for store/folder exclusion, driven by
    /// <see cref="FolderExclusionOptions"/>. Used both by the indexing service
    /// (store-level) and by the Adapter during its walk (folder-level), so the
    /// rules are defined and unit-tested in one place.
    /// </summary>
    public sealed class FolderExclusionPolicy
    {
        private readonly FolderExclusionOptions _options;

        public FolderExclusionPolicy(FolderExclusionOptions options)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
        }

        /// <summary>True if the given store should not be indexed.</summary>
        public bool IsStoreExcluded(StoreInfo store)
        {
            if (store == null) throw new ArgumentNullException(nameof(store));

            if (_options.ExcludePublicFolderStores && store.IsPublicFolderStore)
                return true;

            foreach (var substring in _options.ExcludedStoreNameSubstrings)
            {
                if (ContainsIgnoreCase(store.DisplayName, substring))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// True if the given folder (and its subtree) should be pruned. The
        /// caller supplies the <paramref name="kind"/> it derived from the OM.
        /// </summary>
        public bool IsFolderExcluded(string folderName, WellKnownFolderKind kind)
        {
            if (_options.ExcludeDeletedItems && kind == WellKnownFolderKind.DeletedItems)
                return true;

            return false;
        }

        private static bool ContainsIgnoreCase(string haystack, string needle)
        {
            if (string.IsNullOrEmpty(haystack) || string.IsNullOrEmpty(needle))
                return false;

            return CultureInfo.InvariantCulture.CompareInfo.IndexOf(
                haystack, needle, CompareOptions.IgnoreCase) >= 0;
        }
    }
}
