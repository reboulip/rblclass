using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using RBLclass.Core;
using Serilog;
// Aliased as OutlookOM (not "Outlook") because this project's own namespace is
// RBLclass.Outlook, which would shadow a plain "Outlook" alias.
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.Outlook.Adapter
{
    /// <summary>
    /// <see cref="IMailStore"/> over the live Outlook Object Model. This is the
    /// reimplementation of the legacy <c>archiveExplo</c>/<c>inDepthExplo</c>
    /// recursive walk, but COM-correct: every store/folder/collection is wrapped
    /// in a <see cref="ComRef{T}"/> and released, loops are index-based, and no
    /// COM property chains are used (CLAUDE.md COM-lifetime rules).
    /// </summary>
    /// <remarks>
    /// All methods touch COM and MUST be called on the Outlook UI (STA) thread.
    /// Excluded by design from automated tests (needs a live Outlook); validated
    /// via the install-and-load loop.
    /// </remarks>
    public sealed class OutlookMailStore : IMailStore
    {
        private readonly OutlookOM.Application _app;
        private readonly FolderExclusionPolicy _policy;

        public OutlookMailStore(OutlookOM.Application app,
                                FolderExclusionOptions exclusionOptions = null)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _policy = new FolderExclusionPolicy(
                exclusionOptions ?? FolderExclusionOptions.Default);
        }

        public IReadOnlyList<StoreInfo> GetStores()
        {
            var list = new List<StoreInfo>();

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            using (var stores = new ComRef<OutlookOM.Stores>(session.Value.Stores))
            {
                int count = stores.Value.Count;
                for (int i = 1; i <= count; i++)
                {
                    OutlookOM.Store raw;
                    try { raw = stores.Value[i]; }
                    catch { continue; } // store offline / cannot open - skip it

                    using (var store = new ComRef<OutlookOM.Store>(raw))
                    {
                        string id = Safe(() => store.Value.StoreID, null);
                        if (id == null) continue;
                        list.Add(new StoreInfo(
                            id,
                            Safe(() => store.Value.DisplayName, string.Empty),
                            Safe(() => store.Value.IsDataFileStore, false),
                            // ExchangeStoreType throws for non-Exchange (IMAP/POP)
                            // stores; Safe() maps that to "not a public folder".
                            Safe(() => store.Value.ExchangeStoreType
                                       == OutlookOM.OlExchangeStoreType.olExchangePublicFolder,
                                 false)));
                    }
                }
            }

            return list;
        }

        public IReadOnlyList<FolderNode> GetFolders(string storeId)
        {
            if (storeId == null) throw new ArgumentNullException(nameof(storeId));

            var result = new List<FolderNode>();

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                OutlookOM.Store rawStore;
                try { rawStore = session.Value.GetStoreFromID(storeId); }
                catch { return result; } // store no longer open - tolerate misses

                using (var store = new ComRef<OutlookOM.Store>(rawStore))
                {
                    string deletedItemsEntryId = SafeDeletedItemsEntryId(store.Value);

                    OutlookOM.Folder rawRoot;
                    try { rawRoot = (OutlookOM.Folder)store.Value.GetRootFolder(); }
                    catch { return result; }

                    using (var root = new ComRef<OutlookOM.Folder>(rawRoot))
                    {
                        string rootName = Safe(() => root.Value.Name, string.Empty);
                        WalkChildren(root.Value, storeId,
                                     parentEntryId: null, parentPath: rootName,
                                     deletedItemsEntryId, result);
                    }
                }
            }

            return result;
        }

        public void NavigateTo(string storeId, string entryId, bool newWindow)
        {
            if (storeId == null || entryId == null) return;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                OutlookOM.Folder rawFolder;
                try
                {
                    rawFolder = (OutlookOM.Folder)session.Value.GetFolderFromID(entryId, storeId);
                }
                catch
                {
                    return; // (StoreId, EntryId) no longer resolves - tolerate the miss
                }

                using (var folder = new ComRef<OutlookOM.Folder>(rawFolder))
                {
                    if (newWindow)
                    {
                        folder.Value.Display();
                        return;
                    }

                    OutlookOM.Explorer rawExplorer;
                    try { rawExplorer = _app.ActiveExplorer(); }
                    catch { rawExplorer = null; }

                    if (rawExplorer == null)
                    {
                        // No active explorer to retarget - fall back to a new window.
                        folder.Value.Display();
                        return;
                    }

                    using (var explorer = new ComRef<OutlookOM.Explorer>(rawExplorer))
                    {
                        explorer.Value.CurrentFolder = folder.Value;
                    }
                }
            }
        }

        public IReadOnlyList<MailItemRef> GetSelectedItems()
        {
            var list = new List<MailItemRef>();

            OutlookOM.Explorer rawExplorer;
            try { rawExplorer = _app.ActiveExplorer(); }
            catch { rawExplorer = null; }
            if (rawExplorer == null) return list;

            using (var explorer = new ComRef<OutlookOM.Explorer>(rawExplorer))
            {
                // Read the selection first (before touching CurrentFolder).
                OutlookOM.Selection rawSelection;
                try { rawSelection = explorer.Value.Selection; }
                catch { return list; }

                using (var selection = new ComRef<OutlookOM.Selection>(rawSelection))
                {
                    int count = 0;
                    try { count = selection.Value.Count; } catch { }

                    for (int i = 1; i <= count; i++)
                    {
                        object raw;
                        try { raw = selection.Value[i]; }
                        catch { continue; }

                        using (var comItem = new ComRef<object>(raw))
                        {
                            // Selection can hold MeetingItem/ReportItem/... - skip non-mail.
                            var mail = comItem.Value as OutlookOM.MailItem;
                            if (mail == null) continue;

                            string entryId = Safe(() => mail.EntryID, null);
                            if (entryId == null) continue;

                            string storeId = GetItemStoreId(mail);
                            if (storeId == null) continue;

                            list.Add(new MailItemRef(storeId, entryId,
                                                     Safe(() => mail.Subject, string.Empty)));
                        }
                    }

                    Log.Information(
                        "GetSelectedItems: Selection.Count={Count}, mail items returned={Returned}.",
                        count, list.Count);
                }
            }

            return list;
        }

        public FolderNode CreateSubfolder(FolderNode parent, string name)
        {
            if (parent == null || string.IsNullOrWhiteSpace(name)) return null;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                OutlookOM.Folder rawParent;
                try { rawParent = (OutlookOM.Folder)session.Value.GetFolderFromID(
                    parent.EntryId, parent.StoreId); }
                catch { return null; }

                using (var parentFolder = new ComRef<OutlookOM.Folder>(rawParent))
                using (var folders = new ComRef<OutlookOM.Folders>(parentFolder.Value.Folders))
                {
                    OutlookOM.Folder rawNew;
                    try { rawNew = (OutlookOM.Folder)folders.Value.Add(name, Type.Missing); }
                    catch { return null; }

                    using (var newFolder = new ComRef<OutlookOM.Folder>(rawNew))
                    {
                        string entryId = Safe(() => newFolder.Value.EntryID, null);
                        if (entryId == null) return null;
                        string fullPath = parent.FullPath + FolderNode.PathSeparator + name;
                        return new FolderNode(parent.StoreId, entryId, parent.EntryId,
                                              name, fullPath, isLeaf: true);
                    }
                }
            }
        }

        /// <summary>Store id of the folder containing a mail item (via its parent).</summary>
        private static string GetItemStoreId(OutlookOM.MailItem mail)
        {
            ComRef<OutlookOM.Folder> parent = null;
            try
            {
                parent = new ComRef<OutlookOM.Folder>((OutlookOM.Folder)mail.Parent);
                return Safe(() => parent.Value.StoreID, null);
            }
            catch { return null; }
            finally { parent?.Dispose(); }
        }

        public MailItemRef CopyItemToFolder(MailItemRef item, FolderNode destination)
        {
            if (item == null || destination == null) return null;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return null; } // item no longer resolves - tolerate the miss

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return null;

                    OutlookOM.Folder rawDest;
                    try { rawDest = (OutlookOM.Folder)session.Value.GetFolderFromID(
                        destination.EntryId, destination.StoreId); }
                    catch { return null; }

                    using (var destFolder = new ComRef<OutlookOM.Folder>(rawDest))
                    using (var comCopy = new ComRef<object>(mail.Copy()))
                    {
                        var copy = comCopy.Value as OutlookOM.MailItem;
                        if (copy == null) return null;

                        // Move returns the moved item; the copy reference is now
                        // invalid (CLAUDE.md). Read the moved item's id for the
                        // returned ref, then release it.
                        using (var comMoved = new ComRef<object>(copy.Move(destFolder.Value)))
                        {
                            var moved = comMoved.Value as OutlookOM.MailItem;
                            if (moved == null) return null;
                            string movedEntryId = Safe(() => moved.EntryID, null);
                            if (movedEntryId == null) return null;
                            return new MailItemRef(destination.StoreId, movedEntryId, item.Subject);
                        }
                    }
                }
            }
        }

        public void DeleteItem(MailItemRef item)
        {
            if (item == null) return;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail != null)
                    {
                        try { mail.Delete(); } catch { }
                    }
                }
            }
        }

        public IReadOnlyList<MailItemRef> GetConversationSiblings(MailItemRef item)
        {
            var result = new List<MailItemRef>();
            if (item == null) return result;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return result; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return result;

                    OutlookOM.Conversation rawConversation;
                    try { rawConversation = mail.GetConversation(); }
                    catch { return result; } // conversation tracking unavailable - tolerate

                    if (rawConversation == null) return result;

                    using (var conversation = new ComRef<OutlookOM.Conversation>(rawConversation))
                    {
                        // Scope to the default store's Inbox + Sent Items, the
                        // legacy 5b-step-2 scope. GetTable() gives Outlook's own
                        // flattened, indexed view of every item in the
                        // conversation (anywhere) - resolving and filtering that
                        // small set is far cheaper and more robust than the
                        // legacy approach of iterating every item in both
                        // folders and string-comparing ConversationID.
                        string inboxEntryId = SafeNamespaceDefaultFolderEntryId(
                            session.Value, OutlookOM.OlDefaultFolders.olFolderInbox);
                        string sentEntryId = SafeNamespaceDefaultFolderEntryId(
                            session.Value, OutlookOM.OlDefaultFolders.olFolderSentMail);
                        if (inboxEntryId == null && sentEntryId == null) return result;

                        OutlookOM.Table rawTable;
                        try { rawTable = conversation.Value.GetTable(); }
                        catch { return result; }

                        using (var table = new ComRef<OutlookOM.Table>(rawTable))
                        {
                            var seen = new HashSet<string> { item.EntryId };

                            table.Value.MoveToStart();
                            while (!table.Value.EndOfTable)
                            {
                                OutlookOM.Row row;
                                try { row = table.Value.GetNextRow(); }
                                catch { break; }

                                string entryId = Safe(() => row["EntryID"] as string, null);
                                if (entryId == null || !seen.Add(entryId)) continue;

                                AddSiblingIfInScope(session.Value, entryId,
                                                     inboxEntryId, sentEntryId, result);
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Resolve <paramref name="entryId"/> and, if it is a plain mail item
        /// living directly in the Inbox or Sent Items (by folder EntryID, and
        /// not S/MIME-signed/encrypted - legacy skips <c>IPM.Note.SMIME</c>),
        /// append it to <paramref name="result"/>.
        /// </summary>
        private static void AddSiblingIfInScope(OutlookOM.NameSpace session, string entryId,
                                                 string inboxEntryId, string sentEntryId,
                                                 List<MailItemRef> result)
        {
            object rawSibling;
            try { rawSibling = session.GetItemFromID(entryId); }
            catch { return; }

            using (var comSibling = new ComRef<object>(rawSibling))
            {
                var sibling = comSibling.Value as OutlookOM.MailItem;
                if (sibling == null) return;
                if (Safe(() => sibling.MessageClass, string.Empty) == "IPM.Note.SMIME") return;

                if (!TryGetFolderInfo(sibling, out string storeId, out string folderEntryId))
                    return;
                if (folderEntryId != inboxEntryId && folderEntryId != sentEntryId) return;

                result.Add(new MailItemRef(storeId, entryId, Safe(() => sibling.Subject, string.Empty)));
            }
        }

        /// <summary>Store id and containing-folder id of a mail item, opening its Parent exactly once.</summary>
        private static bool TryGetFolderInfo(OutlookOM.MailItem mail, out string storeId, out string folderEntryId)
        {
            storeId = null;
            folderEntryId = null;
            ComRef<OutlookOM.Folder> parent = null;
            try
            {
                parent = new ComRef<OutlookOM.Folder>((OutlookOM.Folder)mail.Parent);
                storeId = Safe(() => parent.Value.StoreID, null);
                folderEntryId = Safe(() => parent.Value.EntryID, null);
                return storeId != null && folderEntryId != null;
            }
            catch { return false; }
            finally { parent?.Dispose(); }
        }

        private static string SafeNamespaceDefaultFolderEntryId(OutlookOM.NameSpace session,
                                                                 OutlookOM.OlDefaultFolders kind)
        {
            ComRef<OutlookOM.Folder> folder = null;
            try
            {
                folder = new ComRef<OutlookOM.Folder>((OutlookOM.Folder)session.GetDefaultFolder(kind));
                return Safe(() => folder.Value.EntryID, null);
            }
            catch { return null; } // no default store / folder unavailable - tolerate
            finally { folder?.Dispose(); }
        }

        public bool IsFlaggedIncomplete(MailItemRef item)
        {
            if (item == null) return false;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return false; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return false;

                    // Real OM semantics, not the legacy locale-formatted
                    // "01/01/4501" TaskCompletedDate sentinel: a follow-up flag
                    // that hasn't been completed reports olFlagMarked (vs
                    // olFlagComplete / olNoFlag).
                    return Safe(() => mail.IsMarkedAsTask, false)
                        && Safe(() => mail.FlagStatus, OutlookOM.OlFlagStatus.olNoFlag)
                               == OutlookOM.OlFlagStatus.olFlagMarked;
                }
            }
        }

        public void MarkTaskComplete(MailItemRef item)
        {
            if (item == null) return;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return;

                    try
                    {
                        mail.FlagStatus = OutlookOM.OlFlagStatus.olFlagComplete;
                        mail.Save();
                    }
                    catch { }
                }
            }
        }

        public void RemoveAttachments(MailItemRef item)
        {
            if (item == null) return;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return;

                    using (var attachments = new ComRef<OutlookOM.Attachments>(mail.Attachments))
                    {
                        int count = attachments.Value.Count;
                        bool removedAny = false;
                        // Remove from the end so indexes don't shift underneath us.
                        for (int i = count; i >= 1; i--)
                        {
                            try { attachments.Value.Remove(i); removedAny = true; }
                            catch { }
                        }
                        if (removedAny)
                        {
                            try { mail.Save(); } catch { }
                        }
                    }
                }
            }
        }

        public SendGuardInfo InspectForSend(object item)
        {
            var mail = item as OutlookOM.MailItem;
            if (mail == null) return null;

            string body = Safe(() => mail.Body, string.Empty);

            int attachmentCount;
            using (var attachments = new ComRef<OutlookOM.Attachments>(mail.Attachments))
                attachmentCount = Safe(() => attachments.Value.Count, 0);

            var recipients = new List<RecipientAddress>();
            using (var recipientsRef = new ComRef<OutlookOM.Recipients>(mail.Recipients))
            {
                int count = Safe(() => recipientsRef.Value.Count, 0);
                for (int i = 1; i <= count; i++)
                {
                    OutlookOM.Recipient raw;
                    try { raw = recipientsRef.Value[i]; }
                    catch { continue; }

                    using (var recipient = new ComRef<OutlookOM.Recipient>(raw))
                        recipients.Add(ReadRecipient(recipient.Value));
                }
            }

            return new SendGuardInfo(body, attachmentCount, recipients);
        }

        private static RecipientAddress ReadRecipient(OutlookOM.Recipient recipient)
        {
            string name = Safe(() => recipient.Name, null);
            string address = Safe(() => recipient.Address, null);
            bool exchangeResolved = false;

            try
            {
                using (var addressEntry = new ComRef<OutlookOM.AddressEntry>(recipient.AddressEntry))
                {
                    if (addressEntry.Value != null)
                    {
                        var userType = Safe(() => addressEntry.Value.AddressEntryUserType,
                                             OutlookOM.OlAddressEntryUserType.olOtherAddressEntry);
                        exchangeResolved = userType.ToString().StartsWith("olExchange", StringComparison.Ordinal);
                    }
                }
            }
            catch { }

            return new RecipientAddress(name, address, exchangeResolved);
        }

        /// <summary>
        /// Enumerate the immediate sub-folders of <paramref name="parent"/> and
        /// index each (recursing into its subtree). Opens <c>parent.Folders</c>
        /// exactly once.
        /// </summary>
        private void WalkChildren(OutlookOM.Folder parent,
                                  string storeId,
                                  string parentEntryId,
                                  string parentPath,
                                  string deletedItemsEntryId,
                                  List<FolderNode> result)
        {
            using (var subs = new ComRef<OutlookOM.Folders>(parent.Folders))
            {
                int count = subs.Value.Count;
                for (int i = 1; i <= count; i++)
                {
                    OutlookOM.Folder raw;
                    try { raw = (OutlookOM.Folder)subs.Value[i]; }
                    catch { continue; }

                    using (var child = new ComRef<OutlookOM.Folder>(raw))
                    {
                        string entryId = Safe(() => child.Value.EntryID, null);
                        if (entryId == null) continue;

                        string name = Safe(() => child.Value.Name, string.Empty);

                        var kind = (deletedItemsEntryId != null && entryId == deletedItemsEntryId)
                            ? WellKnownFolderKind.DeletedItems
                            : WellKnownFolderKind.Normal;
                        if (_policy.IsFolderExcluded(name, kind))
                            continue; // prune the whole subtree

                        string fullPath = parentPath + FolderNode.PathSeparator + name;

                        // Open this child's sub-folders once: the count tells us
                        // leafness, and we reuse the recursion to descend.
                        int subCount = SafeSubfolderCount(child.Value);
                        result.Add(new FolderNode(storeId, entryId, parentEntryId,
                                                  name, fullPath, isLeaf: subCount == 0));

                        if (subCount > 0)
                            WalkChildren(child.Value, storeId, entryId, fullPath,
                                         deletedItemsEntryId, result);
                    }
                }
            }
        }

        private static int SafeSubfolderCount(OutlookOM.Folder folder)
        {
            ComRef<OutlookOM.Folders> subs = null;
            try
            {
                subs = new ComRef<OutlookOM.Folders>(folder.Folders);
                return subs.Value.Count;
            }
            catch { return 0; }
            finally { subs?.Dispose(); }
        }

        private static string SafeDeletedItemsEntryId(OutlookOM.Store store)
        {
            ComRef<OutlookOM.Folder> deleted = null;
            try
            {
                deleted = new ComRef<OutlookOM.Folder>(
                    (OutlookOM.Folder)store.GetDefaultFolder(
                        OutlookOM.OlDefaultFolders.olFolderDeletedItems));
                return deleted.Value.EntryID;
            }
            catch { return null; } // store has no Deleted Items default - fine
            finally { deleted?.Dispose(); }
        }

        private static T Safe<T>(Func<T> read, T fallback)
        {
            try { return read(); }
            catch { return fallback; }
        }
    }
}
