using System;

namespace RBLclass.Core
{
    /// <summary>
    /// One node in the folder tree (folders only - mails are never indexed,
    /// matching legacy behaviour). Replaces the legacy <c>FolderInPST</c>, but
    /// keyed by the stable <see cref="StoreId"/> + <see cref="EntryId"/> pair
    /// rather than fragile name strings (CLAUDE.md: "cache by (StoreID,
    /// EntryID), tolerate misses").
    /// </summary>
    public sealed class FolderNode
    {
        /// <summary>
        /// Separator used to build <see cref="FullPath"/> from name segments.
        /// Shared by the Adapter (which builds paths during the walk) and the
        /// folder search (Step 2) so both agree on the path shape.
        /// </summary>
        public const string PathSeparator = " / ";

        public FolderNode(string storeId,
                          string entryId,
                          string parentEntryId,
                          string name,
                          string fullPath,
                          bool isLeaf)
        {
            StoreId = storeId ?? throw new ArgumentNullException(nameof(storeId));
            EntryId = entryId ?? throw new ArgumentNullException(nameof(entryId));
            ParentEntryId = parentEntryId; // null for store-top-level folders
            Name = name ?? throw new ArgumentNullException(nameof(name));
            FullPath = fullPath ?? throw new ArgumentNullException(nameof(fullPath));
            IsLeaf = isLeaf;
        }

        /// <summary>Identifier of the store (PST) this folder belongs to.</summary>
        public string StoreId { get; }

        /// <summary>Outlook EntryID of this folder. Stable within a session.</summary>
        public string EntryId { get; }

        /// <summary>
        /// EntryID of the parent folder, or <c>null</c> when this folder sits
        /// directly under the store root.
        /// </summary>
        public string ParentEntryId { get; }

        /// <summary>Display name of this folder (leaf segment).</summary>
        public string Name { get; }

        /// <summary>
        /// Human-readable path from the store root to this folder, joined by
        /// <see cref="PathSeparator"/> (e.g. "Archive2024 / Projects / ContractX").
        /// This is the legacy <c>up</c> string, used as the folder-search surface.
        /// </summary>
        public string FullPath { get; }

        /// <summary>True when this folder has no sub-folders.</summary>
        public bool IsLeaf { get; }
    }
}
