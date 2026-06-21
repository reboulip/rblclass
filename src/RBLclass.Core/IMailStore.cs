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

        /// <summary>
        /// Move <paramref name="item"/> into <paramref name="destination"/> and
        /// return a reference to it at its new location, or null when the item
        /// no longer resolves. Unlike <see cref="CopyItemToFolder"/> this
        /// creates no transient copy and leaves nothing behind - since v2.2 it
        /// is how classify files the original when "keep a copy" is off (the
        /// old copy-then-delete dance made the Stormshield add-in chase
        /// vanishing items and throw MAPI_E_NOT_FOUND). Honour the rule that
        /// Move invalidates the moved reference.
        /// </summary>
        MailItemRef MoveItemToFolder(MailItemRef item, FolderNode destination);

        /// <summary>Delete a mail item (sent-item triage "Delete"; classify no longer deletes).</summary>
        void DeleteItem(MailItemRef item);

        /// <summary>
        /// Resolve a store's Deleted Items folder as a filing destination (the
        /// v2.2 opt-in classify safety copy), or null when the store has none /
        /// no longer resolves. Touches COM - call on the Outlook UI thread.
        /// </summary>
        FolderNode GetDeletedItemsFolder(string storeId);

        /// <summary>
        /// The folder currently containing <paramref name="item"/>, or null
        /// when it no longer resolves. Recorded before a classify move so Undo
        /// knows where to put the item back (v2.2).
        /// </summary>
        FolderNode GetParentFolder(MailItemRef item);

        /// <summary>
        /// A stable identifier of the item's Outlook conversation
        /// (<c>MailItem.ConversationID</c>), or null/empty when the item no
        /// longer resolves or the store doesn't track conversations. Keys the
        /// classification history behind Auto-class (v2.2). Touches COM - call
        /// on the Outlook UI thread.
        /// </summary>
        string GetConversationKey(MailItemRef item);

        /// <summary>
        /// Permanently delete a mail item (v2.2 Undo removing the copies a
        /// classify created - "completely", not into Deleted Items).
        /// </summary>
        void DeleteItemPermanently(MailItemRef item);

        /// <summary>
        /// Re-mark an item's follow-up flag as incomplete
        /// (<c>OlFlagStatus.olFlagMarked</c>) and save it - the v2.2 Undo
        /// counterpart of <see cref="MarkTaskComplete"/>.
        /// </summary>
        void MarkTaskIncomplete(MailItemRef item);

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

        /// <summary>
        /// Strip all attachments from a mail item and save it - unless the item
        /// is S/MIME-encrypted/signed (<c>IPM.Note.SMIME*</c>), whose
        /// "attachments" are the message payload itself: those are never
        /// stripped (maintainer rule, 2026-06-12) and the method returns false
        /// so the caller can tell the user. Returns true otherwise.
        /// </summary>
        bool RemoveAttachments(MailItemRef item);

        /// <summary>
        /// List an item's attachments as Outlook-free descriptors (v2.4.0.0 F2),
        /// for the disposition modal. Returns an empty list when the item does
        /// not resolve or carries none. Touches COM - call on the Outlook UI
        /// thread.
        /// </summary>
        IReadOnlyList<AttachmentInfo> GetAttachments(MailItemRef item);

        /// <summary>
        /// Save one attachment (identified by <paramref name="attachmentId"/> from
        /// <see cref="GetAttachments"/>) into <paramref name="destinationDirectory"/>,
        /// resolving filename collisions. Returns true on success (v2.4.0.0 F2).
        /// Call before stripping, while the attachment still exists. Touches COM -
        /// call on the Outlook UI thread.
        /// </summary>
        bool SaveAttachmentToFile(MailItemRef item, int attachmentId, string destinationDirectory);

        /// <summary>
        /// True when the item is S/MIME-encrypted/signed (<c>IPM.Note.SMIME*</c>),
        /// whose attachments must never be exposed or stripped (v2.4.0.0 F2 uses
        /// this to exclude such mail from the disposition modal and report it).
        /// Touches COM - call on the Outlook UI thread.
        /// </summary>
        bool IsEncryptedMail(MailItemRef item);

        /// <summary>
        /// Remove the learned external-sender banner (<paramref name="bannerSignature"/>,
        /// verbatim block HTML) from a mail item's HTML body and save it, via
        /// <see cref="ExternalBannerStripper"/> (v2.2). No-op (returns false) when
        /// the banner isn't present, the signature is empty, or the item is
        /// S/MIME-encrypted (its body must not be rewritten). Returns true when
        /// the body was changed. Touches COM - call on the Outlook UI thread.
        /// </summary>
        bool StripExternalBanner(MailItemRef item, string bannerSignature);

        /// <summary>
        /// The HTML body of the first selected mail (v2.2 "learn banner" capture),
        /// or null when nothing suitable is selected. The add-in derives the
        /// banner signature from it via <see cref="ExternalBannerStripper.ExtractBannerBlock"/>.
        /// Touches COM - call on the Outlook UI thread.
        /// </summary>
        string GetSelectedItemHtmlBody();

        /// <summary>
        /// Strip the learned banner from a live, possibly-unsaved mail object
        /// (the reply/forward draft Outlook hands us via a new inspector, v2.2),
        /// boxed as <c>object</c> to keep this interface Outlook-free like
        /// <see cref="InspectForSend"/>. Skips encrypted items; returns true when
        /// the body changed. Touches COM - call on the Outlook UI thread.
        /// </summary>
        bool StripBannerFromDraft(object draft, string bannerSignature);

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
