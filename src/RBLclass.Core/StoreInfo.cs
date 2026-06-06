using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Identity of an open Outlook store (PST archive, or Exchange/public
    /// store). Returned by <see cref="IMailStore.GetStores"/>; the indexing
    /// service decides which stores to walk via <see cref="FolderExclusionPolicy"/>.
    /// </summary>
    public sealed class StoreInfo
    {
        public StoreInfo(string storeId,
                         string displayName,
                         bool isDataFileStore,
                         bool isPublicFolderStore = false)
        {
            StoreId = storeId ?? throw new ArgumentNullException(nameof(storeId));
            DisplayName = displayName ?? string.Empty;
            IsDataFileStore = isDataFileStore;
            IsPublicFolderStore = isPublicFolderStore;
        }

        /// <summary>Stable store identifier (Outlook StoreID).</summary>
        public string StoreId { get; }

        /// <summary>Display name of the store.</summary>
        public string DisplayName { get; }

        /// <summary>
        /// True for on-disk .pst data files. False for the primary mailbox
        /// (Exchange/IMAP/OST) and public-folder stores. Informational only -
        /// the index walks the mailbox too (matching the legacy tool), so this
        /// is NOT used to exclude stores; see <see cref="IsPublicFolderStore"/>.
        /// </summary>
        public bool IsDataFileStore { get; }

        /// <summary>
        /// True for an Exchange public-folder store. This is the robust,
        /// locale-free signal used to skip public folders (deviation #8 -
        /// replaces the legacy "Dossiers publics"/"G-FRA" name substrings),
        /// derived from the Outlook store's Exchange store type.
        /// </summary>
        public bool IsPublicFolderStore { get; }
    }
}
