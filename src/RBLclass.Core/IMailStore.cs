using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// The Outlook seam: the only door from the business core to live mail
    /// data. Implemented in RBLclass.Outlook.Adapter over the Outlook Object
    /// Model; faked in tests. Contains no Outlook types so the core stays
    /// portable.
    /// </summary>
    /// <remarks>
    /// Implementations touch COM and therefore must be called on the Outlook
    /// UI (STA) thread (CLAUDE.md threading rules). The interface itself is
    /// thread-agnostic, which is what lets it be driven freely from unit tests.
    /// </remarks>
    public interface IMailStore
    {
        /// <summary>Enumerate all currently open stores.</summary>
        IReadOnlyList<StoreInfo> GetStores();

        /// <summary>
        /// Walk one store and return its folders as a flat list, each carrying
        /// its <see cref="FolderNode.ParentEntryId"/>. Folder-level exclusions
        /// (e.g. Deleted Items) are applied by the implementation during the
        /// walk so excluded subtrees are pruned, not merely hidden.
        /// </summary>
        IReadOnlyList<FolderNode> GetFolders(string storeId);

        /// <summary>
        /// Navigate Outlook to the folder identified by (storeId, entryId),
        /// re-resolving by those stable IDs and tolerating misses (the EntryID
        /// may have changed). Opens a new explorer window when
        /// <paramref name="newWindow"/> is true, otherwise re-targets the active
        /// explorer. Touches COM - call on the Outlook UI thread.
        /// </summary>
        void NavigateTo(string storeId, string entryId, bool newWindow);

        /// <summary>
        /// The mail items currently selected in the active explorer (non-mail
        /// items - meetings, reports - are skipped). Empty if no explorer.
        /// </summary>
        IReadOnlyList<MailItemRef> GetSelectedItems();

        /// <summary>
        /// Place a copy of <paramref name="item"/> into <paramref name="destination"/>
        /// (Outlook Copy-then-Move; the original is untouched) and return a
        /// reference to the filed copy, or null on failure. Honour the rule that
        /// Move invalidates the moved reference.
        /// </summary>
        MailItemRef CopyItemToFolder(MailItemRef item, FolderNode destination);

        /// <summary>Delete a mail item (used to remove originals when not keeping a copy).</summary>
        void DeleteItem(MailItemRef item);

        /// <summary>Strip all attachments from a mail item and save it.</summary>
        void RemoveAttachments(MailItemRef item);

        /// <summary>
        /// Create a sub-folder under <paramref name="parent"/> and return the new
        /// node, or null on failure. Touches COM - call on the Outlook UI thread.
        /// </summary>
        FolderNode CreateSubfolder(FolderNode parent, string name);
    }
}
