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

        /// <summary>
        /// Items sharing <paramref name="item"/>'s conversation, restricted to
        /// the default store's Inbox and Sent Items and excluding
        /// <paramref name="item"/> itself (legacy "conversation widening",
        /// 5b step 2). Encrypted/signed (S/MIME) siblings can't be safely
        /// processed when the provider (e.g. Stormshield) is inactive, so they
        /// are reported separately (<see cref="ConversationSiblings.SkippedEncryptedSubjects"/>)
        /// rather than included. Returns <see cref="ConversationSiblings.Empty"/>
        /// when the item has no in-scope conversation.
        /// </summary>
        ConversationSiblings GetConversationSiblings(MailItemRef item);

        /// <summary>
        /// True when the item carries a follow-up flag that has not been marked
        /// complete (legacy "task-completion guard" check, 5b step 3 - done via
        /// <c>FlagStatus</c>/<c>IsMarkedAsTask</c>, never a locale-formatted date
        /// sentinel).
        /// </summary>
        bool IsFlaggedIncomplete(MailItemRef item);

        /// <summary>Mark an item's follow-up flag complete (<c>OlFlagStatus.olFlagComplete</c>) and save it.</summary>
        void MarkTaskComplete(MailItemRef item);

        /// <summary>Strip all attachments from a mail item and save it.</summary>
        void RemoveAttachments(MailItemRef item);

        /// <summary>
        /// Create a sub-folder under <paramref name="parent"/> and return the new
        /// node, or null on failure. Touches COM - call on the Outlook UI thread.
        /// </summary>
        FolderNode CreateSubfolder(FolderNode parent, string name);

        /// <summary>
        /// Extract what the send-time guards (legacy 6a-6b) need - body text,
        /// attachment count, recipient addresses - from a live mail item that's
        /// about to go out. <paramref name="item"/> is the raw <c>object</c>
        /// Outlook hands <c>Application.ItemSend</c>, boxed to keep this
        /// interface Outlook-free; returns null for anything that isn't a
        /// <c>MailItem</c> (meeting requests etc. - the legacy guard is
        /// olMail-only too). Touches COM - call on the Outlook UI thread.
        /// </summary>
        SendGuardInfo InspectForSend(object item);

        /// <summary>
        /// Build a stable <see cref="MailItemRef"/> for a live item Outlook just
        /// handed us (e.g. the Sent Items <c>Items.ItemAdd</c> event, legacy
        /// 6c). <paramref name="item"/> is boxed <c>object</c> to keep this
        /// interface Outlook-free, mirroring <see cref="InspectForSend"/>;
        /// returns null for anything that isn't a <c>MailItem</c>. Touches COM -
        /// call on the Outlook UI thread.
        /// </summary>
        MailItemRef ResolveMailItem(object item);

        /// <summary>
        /// Resolve the default store's Inbox as a filing destination (legacy 6c
        /// "move to Inbox"), or null when unavailable. Resolved by
        /// <c>OlDefaultFolders.olFolderInbox</c> - locale-independent, unlike
        /// matching on the display name. Touches COM - call on the Outlook UI
        /// thread.
        /// </summary>
        FolderNode GetInboxFolder();
    }
}
