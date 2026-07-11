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

        // True after the first conversation-table timeout of the session
        // (0x...4005 - Windows Search / conversation index unavailable for the
        // active stores, ~30s stall per call). Subsequent GetConversationSiblings
        // calls skip GetTable() and go straight to the MAPI folder scan, which
        // needs no index (B3, v2.9.0.0). Instance-lifetime by design: resets on
        // Outlook restart, when the index may have become available again.
        private bool _wideningUnavailable;

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

                        // Index the store root itself as a fileable destination so
                        // users can classify directly at the top of a PST. It sits
                        // alongside the top-level folders (ParentEntryId null), so it
                        // surfaces when the store name is searched without re-parenting
                        // the rest of the tree. CopyItemToFolder resolves it by EntryID
                        // like any other folder.
                        string rootEntryId = Safe(() => root.Value.EntryID, null);
                        if (rootEntryId != null)
                            result.Add(new FolderNode(storeId, rootEntryId,
                                                      parentEntryId: null, name: rootName,
                                                      fullPath: rootName, isLeaf: false));

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
                        // Activate the explorer after a programmatic folder change so
                        // Outlook fires its internal "explorer activated" event chain.
                        // Without this, the explorer's keyboard handler (Delete, etc.)
                        // remains dormant - the CTP button retained Win32 focus and
                        // Outlook never saw the Explorer.Activate event that arms its
                        // message-list accelerators (B2, v2.9.0.0).
                        try { explorer.Value.Activate(); } catch { }
                    }
                }
            }
        }

        public void SelectItemsInFolder(string storeId, string entryId, IReadOnlyList<MailItemRef> items)
        {
            if (storeId == null || entryId == null || items == null || items.Count == 0) return;

            OutlookOM.Explorer rawExplorer;
            try { rawExplorer = _app.ActiveExplorer(); }
            catch { rawExplorer = null; }
            if (rawExplorer == null) return;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            using (var explorer = new ComRef<OutlookOM.Explorer>(rawExplorer))
            {
                OutlookOM.Folder rawFolder;
                try { rawFolder = (OutlookOM.Folder)session.Value.GetFolderFromID(entryId, storeId); }
                catch { return; } // source folder no longer resolves - tolerate the miss

                using (var folder = new ComRef<OutlookOM.Folder>(rawFolder))
                {
                    try { explorer.Value.CurrentFolder = folder.Value; }
                    catch { return; }
                }

                // Replace whatever was selected with the moved-back items. Both
                // ClearSelection and AddToSelection require the item to live in the
                // now-current folder; a per-item catch tolerates any that don't.
                try { explorer.Value.ClearSelection(); } catch { }

                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    if (item == null) continue;

                    object rawItem;
                    try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                    catch { continue; }

                    using (var comItem = new ComRef<object>(rawItem))
                    {
                        try { explorer.Value.AddToSelection(comItem.Value); } catch { }
                    }
                }
            }
        }

        public IReadOnlyList<MailItemRef> GetSelectedItems(bool includeMeetingItems = false)
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
                            var mail = comItem.Value as OutlookOM.MailItem;
                            if (mail != null)
                            {
                                string entryId = Safe(() => mail.EntryID, null);
                                if (entryId == null) continue;
                                string storeId = GetItemStoreId(mail);
                                if (storeId == null) continue;
                                list.Add(new MailItemRef(storeId, entryId,
                                                         Safe(() => mail.Subject, string.Empty)));
                                continue;
                            }

                            if (includeMeetingItems)
                            {
                                var meeting = comItem.Value as OutlookOM.MeetingItem;
                                if (meeting != null)
                                {
                                    string entryId = Safe(() => meeting.EntryID, null);
                                    if (entryId == null) continue;
                                    string storeId = GetMeetingItemStoreId(meeting);
                                    if (storeId == null) continue;
                                    list.Add(new MailItemRef(storeId, entryId,
                                                             Safe(() => meeting.Subject, string.Empty),
                                                             isMeetingItem: true));
                                }
                            }
                            // ReportItem and other non-mail types are always skipped.
                        }
                    }

                    Log.Debug(
                        "GetSelectedItems: Selection.Count={Count}, items returned={Returned} (meetingItems={IncludeMeeting}).",
                        count, list.Count, includeMeetingItems);
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

        private static string GetMeetingItemStoreId(OutlookOM.MeetingItem meeting)
        {
            ComRef<OutlookOM.Folder> parent = null;
            try
            {
                parent = new ComRef<OutlookOM.Folder>((OutlookOM.Folder)meeting.Parent);
                return Safe(() => parent.Value.StoreID, null);
            }
            catch { return null; }
            finally { parent?.Dispose(); }
        }

        public MailItemRef CopyItemToFolder(MailItemRef item, FolderNode destination, bool markAsRead = false)
        {
            if (item == null || destination == null) return null;

            // PERF (v2.5.1.0): time the whole copy-then-move primitive, including
            // the session open + item/folder resolution, so a classify-time
            // timeline can be reconstructed from the rolling log on the target.
            var swPerf = System.Diagnostics.Stopwatch.StartNew();

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch (Exception ex)
                {
                    // item no longer resolves - tolerate the miss
                    Log.Debug(ex, "CopyItemToFolder: item {EntryId} no longer resolves.", item.EntryId);
                    return null;
                }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    var meeting = mail == null && item.IsMeetingItem
                                  ? comItem.Value as OutlookOM.MeetingItem
                                  : null;
                    if (mail == null && meeting == null)
                    {
                        Log.Debug("CopyItemToFolder: item {EntryId} resolved to neither MailItem nor MeetingItem.", item.EntryId);
                        return null;
                    }

                    OutlookOM.Folder rawDest;
                    try { rawDest = (OutlookOM.Folder)session.Value.GetFolderFromID(
                        destination.EntryId, destination.StoreId); }
                    catch (Exception ex)
                    {
                        Log.Debug(ex, "CopyItemToFolder: destination folder {Path} no longer resolves.", destination.FullPath);
                        return null;
                    }

                    using (var destFolder = new ComRef<OutlookOM.Folder>(rawDest))
                    {
                        // Manually managed (not using) so the catch below can still
                        // reach the transient copy to delete it (v2.7.1.0 A3).
                        ComRef<object> comCopy = null;
                        try
                        {
                            object copyRaw = mail != null ? mail.Copy() : meeting.Copy();
                            comCopy = new ComRef<object>(copyRaw);

                            var copyMail    = comCopy.Value as OutlookOM.MailItem;
                            var copyMeeting = copyMail == null ? comCopy.Value as OutlookOM.MeetingItem : null;
                            if (copyMail == null && copyMeeting == null)
                            {
                                Log.Warning("CopyItemToFolder: Copy() of item {EntryId} returned an unusable object; the copy may be stranded in the source folder.", item.EntryId);
                                return null;
                            }

                            // Mark the copy read before it is placed (v2.7.0.0 D2):
                            // setting UnRead on the not-yet-arrived copy is race-free,
                            // unlike a re-fetch after Move which Outlook's arrival
                            // handling can clobber. S/MIME guard and MeetingItem
                            // support delegated to TryMarkItemRead (v2.7.1.0 A1).
                            if (markAsRead)
                                TryMarkItemRead(copyMail, copyMeeting);

                            // Move returns the moved item; the copy reference is now
                            // invalid (CLAUDE.md). Read the moved item's id for the
                            // returned ref, then release it.
                            object movedRaw = copyMail != null
                                ? copyMail.Move(destFolder.Value)
                                : copyMeeting.Move(destFolder.Value);
                            using (var comMoved = new ComRef<object>(movedRaw))
                            {
                                var movedMail    = comMoved.Value as OutlookOM.MailItem;
                                var movedMeeting = movedMail == null ? comMoved.Value as OutlookOM.MeetingItem : null;
                                string movedEntryId =
                                    movedMail    != null ? Safe(() => movedMail.EntryID, null) :
                                    movedMeeting != null ? Safe(() => movedMeeting.EntryID, null) : null;
                                if (movedEntryId == null)
                                {
                                    Log.Warning(
                                        "CopyItemToFolder: item {EntryId} copied to {Path} but the filed copy's EntryID could not be read; it won't be tracked for Undo/history.",
                                        item.EntryId, destination.FullPath);
                                    return null;
                                }
                                swPerf.Stop();
                                Log.Debug(
                                    "PERF CopyItemToFolder {Ms} ms: item {EntryId} -> {Path}",
                                    swPerf.ElapsedMilliseconds, item.EntryId, destination.FullPath);
                                return new MailItemRef(destination.StoreId, movedEntryId,
                                                       item.Subject, item.IsMeetingItem);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Best-effort: if Copy() succeeded but Move() threw, delete
                            // the transient copy before rethrowing so no orphaned
                            // duplicate accumulates in the source folder (v2.7.1.0 A3).
                            if (comCopy != null)
                            {
                                try
                                {
                                    var orphan = comCopy.Value as OutlookOM.MailItem;
                                    if (orphan != null) orphan.Delete();
                                    else (comCopy.Value as OutlookOM.MeetingItem)?.Delete();
                                }
                                catch (Exception cleanupEx)
                                {
                                    Log.Warning(cleanupEx,
                                        "Orphaned copy of item {EntryId} could not be deleted after a failed move.",
                                        item.EntryId);
                                }
                            }
                            // Some folders (notably a store root) can reject items.
                            // Log the cause and rethrow so the classifier counts this
                            // destination as failed and keeps the original intact.
                            Log.Warning(ex,
                                "Filing into folder {Path} ({EntryId}) failed; it may not accept items.",
                                destination.FullPath, destination.EntryId);
                            throw;
                        }
                        finally
                        {
                            comCopy?.Dispose();
                        }
                    }
                }
            }
        }

        public MailItemRef MoveItemToFolder(MailItemRef item, FolderNode destination, bool markAsRead = false)
        {
            if (item == null || destination == null) return null;

            // PERF (v2.5.1.0): full-method timer (session open + resolution +
            // move). Compared against the move-only D1 timing below, the
            // difference isolates the per-call resolution overhead - the
            // suspected tax behind the slow per-item auto-class loop.
            var swPerf = System.Diagnostics.Stopwatch.StartNew();

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch (Exception ex)
                {
                    Log.Debug(ex, "MoveItemToFolder: item {EntryId} no longer resolves.", item.EntryId);
                    return null;
                }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    var meeting = mail == null && item.IsMeetingItem
                                  ? comItem.Value as OutlookOM.MeetingItem
                                  : null;
                    if (mail == null && meeting == null)
                    {
                        Log.Debug("MoveItemToFolder: item {EntryId} resolved to neither MailItem nor MeetingItem.", item.EntryId);
                        return null;
                    }

                    OutlookOM.Folder rawDest;
                    try { rawDest = (OutlookOM.Folder)session.Value.GetFolderFromID(
                        destination.EntryId, destination.StoreId); }
                    catch (Exception ex)
                    {
                        Log.Debug(ex, "MoveItemToFolder: destination folder {Path} no longer resolves.", destination.FullPath);
                        return null;
                    }

                    using (var destFolder = new ComRef<OutlookOM.Folder>(rawDest))
                    {
                        try
                        {
                            // D1 investigation (v2.4): BEGIN/END timing markers
                            // around the COM move, so the multi-item-classify
                            // MAPI_E_NOT_FOUND from Stormshield's OnBeforeReadAsync
                            // can be correlated against when each item was moving.
                            // Gated behind DiagnosticLogging (v2.8.0.0 logging
                            // audit) - flip it on to get this trace on the target.
                            var sw = System.Diagnostics.Stopwatch.StartNew();
                            Log.Debug(
                                "Classify move BEGIN (D1): item {EntryId} -> {Path}",
                                item.EntryId, destination.FullPath);

                            // Move returns the moved item; the original reference
                            // is now invalid (CLAUDE.md). No transient copy is
                            // created and nothing lands in Deleted Items - the
                            // point of the v2.2 move-based classify (Stormshield
                            // chased the old transient copies into
                            // MAPI_E_NOT_FOUND).
                            object movedRaw = mail != null
                                ? mail.Move(destFolder.Value)
                                : meeting.Move(destFolder.Value);
                            using (var comMoved = new ComRef<object>(movedRaw))
                            {
                                var movedMail    = comMoved.Value as OutlookOM.MailItem;
                                var movedMeeting = movedMail == null ? comMoved.Value as OutlookOM.MeetingItem : null;

                                // Mark read only AFTER a successful move, on the moved
                                // item itself (v2.7.1.0 F2): a failed move must not
                                // leave the mail marked read but unfiled. S/MIME guard
                                // and MeetingItem support via TryMarkItemRead.
                                if (markAsRead)
                                    TryMarkItemRead(movedMail, movedMeeting);

                                string movedEntryId =
                                    movedMail    != null ? Safe(() => movedMail.EntryID, null) :
                                    movedMeeting != null ? Safe(() => movedMeeting.EntryID, null) : null;
                                if (movedEntryId == null)
                                {
                                    Log.Warning(
                                        "MoveItemToFolder: item {EntryId} moved to {Path} but the moved item's EntryID could not be read; it won't be tracked for Undo/history.",
                                        item.EntryId, destination.FullPath);
                                    return null;
                                }
                                sw.Stop();
                                swPerf.Stop();
                                Log.Debug(
                                    "Classify move END (D1): item {EntryId} -> {Path} in {Ms} ms",
                                    item.EntryId, destination.FullPath, sw.ElapsedMilliseconds);
                                Log.Debug(
                                    "PERF MoveItemToFolder {Ms} ms (move-only {MoveMs} ms): item {EntryId} -> {Path}",
                                    swPerf.ElapsedMilliseconds, sw.ElapsedMilliseconds,
                                    item.EntryId, destination.FullPath);
                                return new MailItemRef(destination.StoreId, movedEntryId,
                                                       item.Subject, item.IsMeetingItem);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Some folders (notably a store root) can reject items.
                            // Log and rethrow so the classifier counts the failure;
                            // the original stays where it was.
                            Log.Warning(ex,
                                "Moving into folder {Path} ({EntryId}) failed; it may not accept items.",
                                destination.FullPath, destination.EntryId);
                            throw;
                        }
                    }
                }
            }
        }

        public void DeleteItem(MailItemRef item)
        {
            if (item == null) return;

            Log.Information("DeleteItem: deleting {EntryId}", item.EntryId);

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch (Exception ex)
                {
                    Log.Warning(ex, "DeleteItem: item {EntryId} no longer resolves.", item.EntryId);
                    return;
                }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail != null)
                    {
                        try { mail.Delete(); Log.Information("DeleteItem: deleted {EntryId}", item.EntryId); }
                        catch (Exception ex) { Log.Warning(ex, "DeleteItem: delete failed for {EntryId}", item.EntryId); }
                    }
                    else
                    {
                        var meeting = comItem.Value as OutlookOM.MeetingItem;
                        if (meeting != null)
                        {
                            try { meeting.Delete(); Log.Information("DeleteItem: deleted {EntryId}", item.EntryId); }
                            catch (Exception ex) { Log.Warning(ex, "DeleteItem: delete failed for {EntryId}", item.EntryId); }
                        }
                    }
                }
            }
        }

        public FolderNode GetParentFolder(MailItemRef item)
        {
            if (item == null) return null;

            // PERF (v2.5.1.0): resolved once per item on the !keepCopy classify
            // path (to make a move undoable).
            var swPerf = System.Diagnostics.Stopwatch.StartNew();

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return null; } // item no longer resolves - tolerate the miss

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    var meeting = mail == null && item.IsMeetingItem
                                  ? comItem.Value as OutlookOM.MeetingItem
                                  : null;
                    if (mail == null && meeting == null) return null;

                    ComRef<OutlookOM.Folder> parent = null;
                    try
                    {
                        parent = new ComRef<OutlookOM.Folder>(
                            mail != null
                                ? (OutlookOM.Folder)mail.Parent
                                : (OutlookOM.Folder)meeting.Parent);
                        string storeId = Safe(() => parent.Value.StoreID, null);
                        string entryId = Safe(() => parent.Value.EntryID, null);
                        if (storeId == null || entryId == null) return null;

                        string name = Safe(() => parent.Value.Name, string.Empty);
                        // FolderPath is "\\Store\A\B"; good enough for Undo's
                        // purposes (a destination ref + logging) without a walk.
                        string fullPath = Safe(() => parent.Value.FolderPath, name);
                        swPerf.Stop();
                        Log.Debug("PERF GetParentFolder {Ms} ms: item {EntryId}",
                                        swPerf.ElapsedMilliseconds, item.EntryId);
                        return new FolderNode(storeId, entryId, parentEntryId: null,
                                              name: name, fullPath: fullPath, isLeaf: false);
                    }
                    catch { return null; }
                    finally { parent?.Dispose(); }
                }
            }
        }

        public string GetConversationKey(MailItemRef item)
        {
            if (item == null) return null;

            // PERF (v2.5.1.0): read per item for the Auto-class history (once on a
            // manual classify, again per item inside the auto-class loop).
            var swPerf = System.Diagnostics.Stopwatch.StartNew();

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return null; } // item no longer resolves - tolerate the miss

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return null;

                    // ConversationID is stable across the thread and survives
                    // moves between folders/stores - the right key for the
                    // Auto-class history. Null/empty when conversation tracking
                    // is off for the item's store.
                    string key = Safe(() => mail.ConversationID, null);
                    swPerf.Stop();
                    Log.Debug("PERF GetConversationKey {Ms} ms: item {EntryId}",
                                    swPerf.ElapsedMilliseconds, item.EntryId);
                    return key;
                }
            }
        }

        public void DeleteItemPermanently(MailItemRef item)
        {
            if (item == null) return;

            // The OM has no direct hard delete: Delete() only moves to Deleted
            // Items, under a NEW EntryID we never learn. So: move it to Deleted
            // Items ourselves (keeping the reference), then Delete() there -
            // deleting an item already in Deleted Items removes it for good.
            var deletedItems = GetDeletedItemsFolder(item.StoreId);
            MailItemRef inTrash = deletedItems != null
                ? MoveItemToFolder(item, deletedItems)
                : null;
            if (inTrash == null)
            {
                // No Deleted Items / move failed - fall back to a soft delete.
                DeleteItem(item);
                return;
            }

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(inTrash.EntryId, inTrash.StoreId); }
                catch { return; } // already gone - close enough

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

        public void MarkTaskIncomplete(MailItemRef item)
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
                        mail.FlagStatus = OutlookOM.OlFlagStatus.olFlagMarked;
                        mail.Save();
                    }
                    catch { }
                }
            }
        }

        public FolderNode GetDeletedItemsFolder(string storeId)
        {
            if (storeId == null) return null;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                OutlookOM.Store rawStore;
                try { rawStore = session.Value.GetStoreFromID(storeId); }
                catch { return null; } // store no longer open - tolerate the miss

                using (var store = new ComRef<OutlookOM.Store>(rawStore))
                {
                    ComRef<OutlookOM.Folder> deleted = null;
                    try
                    {
                        deleted = new ComRef<OutlookOM.Folder>(
                            (OutlookOM.Folder)store.Value.GetDefaultFolder(
                                OutlookOM.OlDefaultFolders.olFolderDeletedItems));

                        string entryId = Safe(() => deleted.Value.EntryID, null);
                        if (entryId == null) return null;
                        string name = Safe(() => deleted.Value.Name, "Deleted Items");
                        return new FolderNode(storeId, entryId, parentEntryId: null,
                                              name: name, fullPath: name, isLeaf: false);
                    }
                    catch { return null; } // store has no Deleted Items default - fine
                    finally { deleted?.Dispose(); }
                }
            }
        }

        public ConversationSiblings GetConversationSiblings(MailItemRef item)
        {
            var result = new List<MailItemRef>();
            var skipped = new List<string>();

            // PERF (v2.5.1.0): conversation widening on the manual-classify
            // preflight - GetConversation + GetTable + a per-sibling
            // GetItemFromID. A prime suspect for the random classify-time
            // slowness, since it touches the conversation index/server. Logged
            // on every exit path, with the sibling/skip counts it produced.
            var swPerf = System.Diagnostics.Stopwatch.StartNew();
            bool wideningJustDisabled = false;
            ConversationSiblings Outcome()
            {
                swPerf.Stop();
                Log.Debug(
                    "PERF GetConversationSiblings {Ms} ms: item {EntryId}, {Siblings} sibling(s), {Skipped} skipped, path={Path}",
                    swPerf.ElapsedMilliseconds, item?.EntryId, result.Count, skipped.Count,
                    _wideningUnavailable ? "folder-scan" : "conversation-table");
                return new ConversationSiblings(result, skipped, wideningJustDisabled);
            }

            if (item == null) return Outcome();

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return Outcome(); }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return Outcome();

                    // Scope to the default store's Inbox + Sent Items, the
                    // legacy 5b-step-2 scope.
                    string inboxEntryId = SafeNamespaceDefaultFolderEntryId(
                        session.Value, OutlookOM.OlDefaultFolders.olFolderInbox);
                    string sentEntryId = SafeNamespaceDefaultFolderEntryId(
                        session.Value, OutlookOM.OlDefaultFolders.olFolderSentMail);
                    if (inboxEntryId == null && sentEntryId == null) return Outcome();

                    if (_wideningUnavailable)
                    {
                        // A previous call this session already proved the
                        // conversation table times out on this profile - go
                        // straight to the MAPI folder scan (B3, v2.9.0.0).
                        CollectSiblingsViaFolderScan(session.Value, mail, item.EntryId,
                            inboxEntryId, sentEntryId, result, skipped);
                        return Outcome();
                    }

                    OutlookOM.Conversation rawConversation;
                    try { rawConversation = mail.GetConversation(); }
                    catch { return Outcome(); } // conversation tracking unavailable - tolerate

                    if (rawConversation == null) return Outcome();

                    using (var conversation = new ComRef<OutlookOM.Conversation>(rawConversation))
                    {
                        // GetTable() gives Outlook's own flattened, indexed view
                        // of every item in the conversation (anywhere) -
                        // resolving and filtering that small set is far cheaper
                        // and more robust than iterating every item in both
                        // folders and string-comparing ConversationID... when
                        // the conversation index is available at all (see catch).
                        OutlookOM.Table rawTable;
                        try { rawTable = conversation.Value.GetTable(); }
                        catch { return Outcome(); }

                        using (var table = new ComRef<OutlookOM.Table>(rawTable))
                        {
                            var seen = new HashSet<string> { item.EntryId };

                            try
                            {
                                table.Value.MoveToStart();
                                while (!table.Value.EndOfTable)
                                {
                                    OutlookOM.Row row;
                                    try { row = table.Value.GetNextRow(); }
                                    catch { break; }

                                    string entryId = Safe(() => row["EntryID"] as string, null);
                                    if (entryId == null || !seen.Add(entryId)) continue;

                                    AddSiblingIfInScope(session.Value, entryId,
                                                         inboxEntryId, sentEntryId, result, skipped);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Conversation table enumeration failed (0x...4005 -
                                // conversation index unavailable, e.g. Windows Search
                                // not indexing the PST path; burns the full ~30s
                                // timeout per call with 0 rows). Auto-disable the
                                // table for the rest of the session and redo this
                                // item via the MAPI folder scan, discarding any
                                // partial rows - the scan is authoritative for the
                                // Inbox+Sent scope (B3, v2.9.0.0; supersedes the
                                // v2.8.0.0 C2 log-and-tolerate).
                                _wideningUnavailable = true;
                                wideningJustDisabled = true;
                                Log.Warning(ex,
                                    "GetConversationSiblings: conversation table timeout for item {EntryId}; " +
                                    "widening table auto-disabled for this session, using MAPI folder scan.",
                                    item?.EntryId);
                                result.Clear();
                                skipped.Clear();
                                CollectSiblingsViaFolderScan(session.Value, mail, item.EntryId,
                                    inboxEntryId, sentEntryId, result, skipped);
                            }
                        }
                    }
                }
            }

            return Outcome();
        }

        /// <summary>
        /// Resolve <paramref name="entryId"/> and, if it is a mail item living
        /// directly in the Inbox or Sent Items (by folder EntryID), either append
        /// it to <paramref name="result"/> or - when it is S/MIME-signed/encrypted
        /// (<c>IPM.Note.SMIME</c>, e.g. Stormshield, which can't be safely
        /// processed with the provider inactive) - record its subject in
        /// <paramref name="skipped"/> so the caller can warn instead of silently
        /// leaving it behind.
        /// </summary>
        private static void AddSiblingIfInScope(OutlookOM.NameSpace session, string entryId,
                                                 string inboxEntryId, string sentEntryId,
                                                 List<MailItemRef> result, List<string> skipped)
        {
            object rawSibling;
            try { rawSibling = session.GetItemFromID(entryId); }
            catch { return; }

            using (var comSibling = new ComRef<object>(rawSibling))
            {
                var sibling = comSibling.Value as OutlookOM.MailItem;
                if (sibling == null) return;

                if (!TryGetFolderInfo(sibling, out string storeId, out string folderEntryId))
                    return;
                if (folderEntryId != inboxEntryId && folderEntryId != sentEntryId) return;

                // In scope. Encrypted/signed S/MIME is skipped but reported.
                // StartsWith, like every other S/MIME guard in this file, so
                // subtypes (.MultipartSigned, ...) are caught too (v2.7.1.0 A3).
                if (Safe(() => sibling.MessageClass, string.Empty)
                        .StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase))
                {
                    skipped.Add(Safe(() => sibling.Subject, "(encrypted message)"));
                    return;
                }

                result.Add(new MailItemRef(storeId, entryId, Safe(() => sibling.Subject, string.Empty)));
            }
        }

        /// <summary>
        /// MAPI-native sibling lookup, used when the conversation table is
        /// unavailable (B3, v2.9.0.0): Restrict the Inbox and Sent Items by
        /// PR_CONVERSATION_TOPIC (0x0070001E, a plain MAPI string property - no
        /// Windows Search involved), then verify each candidate by exact
        /// ConversationID match, since two unrelated threads can share a
        /// normalized topic ("Re: Meeting"). Same S/MIME skip-and-report rule
        /// as <see cref="AddSiblingIfInScope"/>.
        /// </summary>
        private static void CollectSiblingsViaFolderScan(
            OutlookOM.NameSpace session,
            OutlookOM.MailItem sourceMail,
            string sourceEntryId,
            string inboxEntryId,
            string sentEntryId,
            List<MailItemRef> result,
            List<string> skipped)
        {
            string convId = Safe(() => sourceMail.ConversationID, null);
            if (string.IsNullOrEmpty(convId)) return;

            string topic = Safe(() => sourceMail.ConversationTopic, null);
            string topicFilter = topic != null
                ? "@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x0070001E\" = '"
                  + topic.Replace("'", "''") + "'"
                : null;

            var seen = new HashSet<string> { sourceEntryId };

            foreach (var folderEntryId in new[] { inboxEntryId, sentEntryId })
            {
                if (folderEntryId == null) continue;

                OutlookOM.Folder rawFolder;
                try { rawFolder = (OutlookOM.Folder)session.GetFolderFromID(folderEntryId); }
                catch { continue; }

                using (var folder = new ComRef<OutlookOM.Folder>(rawFolder))
                {
                    string storeId = Safe(() => folder.Value.StoreID, null);
                    if (storeId == null) continue;

                    OutlookOM.Items rawItems;
                    try { rawItems = folder.Value.Items; }
                    catch { continue; }

                    using (var itemsRef = new ComRef<OutlookOM.Items>(rawItems))
                    {
                        OutlookOM.Items rawRestricted = null;
                        if (topicFilter != null)
                        {
                            try { rawRestricted = itemsRef.Value.Restrict(topicFilter); }
                            catch { rawRestricted = null; } // fall back to full iteration
                        }

                        var restrictedRef = rawRestricted != null
                            ? new ComRef<OutlookOM.Items>(rawRestricted)
                            : null;
                        try
                        {
                            OutlookOM.Items toScan = rawRestricted ?? itemsRef.Value;
                            int count = Safe(() => toScan.Count, 0);
                            for (int i = 1; i <= count; i++)
                            {
                                object rawSibling;
                                try { rawSibling = toScan[i]; }
                                catch { continue; }

                                using (var comSibling = new ComRef<object>(rawSibling))
                                {
                                    var sibling = comSibling.Value as OutlookOM.MailItem;
                                    if (sibling == null) continue;

                                    string sibEntryId = Safe(() => sibling.EntryID, null);
                                    if (sibEntryId == null || !seen.Add(sibEntryId)) continue;

                                    string sibConvId = Safe(() => sibling.ConversationID, null);
                                    if (sibConvId != convId) continue;

                                    if (Safe(() => sibling.MessageClass, string.Empty)
                                            .StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase))
                                    {
                                        skipped.Add(Safe(() => sibling.Subject, "(encrypted message)"));
                                        continue;
                                    }

                                    result.Add(new MailItemRef(storeId, sibEntryId,
                                        Safe(() => sibling.Subject, string.Empty)));
                                }
                            }
                        }
                        finally { restrictedRef?.Dispose(); }
                    }
                }
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

            // PERF (v2.5.1.0): checked once per item during the classify preflight.
            var swPerf = System.Diagnostics.Stopwatch.StartNew();

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
                    bool flagged = Safe(() => mail.IsMarkedAsTask, false)
                        && Safe(() => mail.FlagStatus, OutlookOM.OlFlagStatus.olNoFlag)
                               == OutlookOM.OlFlagStatus.olFlagMarked;
                    swPerf.Stop();
                    Log.Debug("PERF IsFlaggedIncomplete {Ms} ms: item {EntryId}",
                                    swPerf.ElapsedMilliseconds, item.EntryId);
                    return flagged;
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

        public bool RemoveAttachments(MailItemRef item)
        {
            if (item == null) return true;

            // PERF (v2.5.1.0): attachment stripping is COM-heavy (a Save per
            // mail, plus PropertyAccessor probes for inline detection) and a
            // candidate for the slow path when "remove attachments" is on.
            var swPerf = System.Diagnostics.Stopwatch.StartNew();

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return true; } // item no longer resolves - nothing to report

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return true;

                    // Never strip S/MIME-encrypted/signed mail (maintainer rule,
                    // 2026-06-12): its "attachments" are the encrypted/signed
                    // payload itself - removing them destroys the message.
                    string messageClass = Safe(() => mail.MessageClass, string.Empty);
                    if (messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase))
                    {
                        Log.Information(
                            "RemoveAttachments skipped an encrypted/signed item ({MessageClass}).",
                            messageClass);
                        return false;
                    }

                    OutlookOM.Attachments rawAtts;
                    // Guarded like the per-attachment reads below: a COM fault here
                    // must not escape mid-classify (v2.7.1.0 B2).
                    try { rawAtts = mail.Attachments; }
                    catch { return true; } // collection inaccessible - nothing to strip
                    using (var attachments = new ComRef<OutlookOM.Attachments>(rawAtts))
                    {
                        int count = Safe(() => attachments.Value.Count, 0);
                        bool removedAny = false;
                        // Remove from the end so indexes don't shift underneath us.
                        // Inline/embedded images are left in place (v2.5.0.0 B2).
                        for (int i = count; i >= 1; i--)
                        {
                            OutlookOM.Attachment rawAtt;
                            try { rawAtt = attachments.Value[i]; } catch { continue; }
                            using (var attRef = new ComRef<OutlookOM.Attachment>(rawAtt))
                            {
                                if (IsInlineAttachment(attRef.Value)) continue;
                                try { attRef.Value.Delete(); removedAny = true; }
                                catch (Exception ex)
                                {
                                    Log.Warning(ex, "RemoveAttachments: deleting one attachment failed for item {EntryId}.", item.EntryId);
                                }
                            }
                        }
                        if (removedAny)
                        {
                            try { mail.Save(); } catch { }
                        }
                        swPerf.Stop();
                        Log.Debug(
                            "PERF RemoveAttachments {Ms} ms: item {EntryId}, {Count} attachment(s) scanned, removedAny={Removed}",
                            swPerf.ElapsedMilliseconds, item.EntryId, count, removedAny);
                    }
                }
            }

            return true;
        }

        public bool RemoveAttachment(MailItemRef item, int attachmentId)
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

                    // Same S/MIME guard as RemoveAttachments: never touch an
                    // encrypted/signed payload.
                    string messageClass = Safe(() => mail.MessageClass, string.Empty);
                    if (messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase))
                    {
                        Log.Information(
                            "RemoveAttachment skipped an encrypted/signed item ({MessageClass}).",
                            messageClass);
                        return false;
                    }

                    using (var attachments = new ComRef<OutlookOM.Attachments>(mail.Attachments))
                    {
                        int count = attachments.Value.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            OutlookOM.Attachment raw;
                            try { raw = attachments.Value[i]; } catch { continue; }
                            using (var att = new ComRef<OutlookOM.Attachment>(raw))
                            {
                                if (Safe(() => att.Value.Index, -1) != attachmentId) continue;
                                try
                                {
                                    att.Value.Delete();
                                    mail.Save();
                                    Log.Information("Removed attachment {Id} from a filed mail.", attachmentId);
                                    return true;
                                }
                                catch (Exception ex)
                                {
                                    Log.Warning(ex, "Removing attachment {Id} failed.", attachmentId);
                                    return false;
                                }
                            }
                        }
                    }
                }
            }

            Log.Warning("RemoveAttachment: attachment {Id} not found.", attachmentId);
            return false;
        }

        public bool AppendHtmlNote(MailItemRef item, string htmlBlock)
        {
            if (item == null || string.IsNullOrWhiteSpace(htmlBlock)) return false;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return false; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return false;

                    // Never rewrite an encrypted/signed body (same rule as the banner strip).
                    string messageClass = Safe(() => mail.MessageClass, string.Empty);
                    if (messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase))
                    {
                        Log.Information(
                            "AppendHtmlNote skipped an encrypted/signed item ({MessageClass}, {EntryId}).",
                            messageClass, item.EntryId);
                        return false;
                    }

                    string html = Safe(() => mail.HTMLBody, null) ?? string.Empty;
                    try
                    {
                        mail.HTMLBody = InsertAfterBodyTag(html, htmlBlock);
                        mail.Save();
                        Log.Information("Appended attachment-disposition label to filed mail {EntryId}.", item.EntryId);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "Appending the attachment-disposition label failed for {EntryId}.", item.EntryId);
                        return false;
                    }
                }
            }
        }

        /// <summary>Insert <paramref name="block"/> just inside the opening &lt;body&gt; tag, or prepend when there is none.</summary>
        private static string InsertAfterBodyTag(string html, string block)
        {
            if (string.IsNullOrEmpty(html)) return block;
            int bodyStart = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            if (bodyStart >= 0)
            {
                int close = html.IndexOf('>', bodyStart);
                if (close >= 0) return html.Insert(close + 1, block);
            }
            return block + html;
        }

        public bool IsEncryptedMail(MailItemRef item)
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
                    string messageClass = Safe(() => mail.MessageClass, string.Empty);
                    return messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase);
                }
            }
        }

        public IReadOnlyList<AttachmentInfo> GetAttachments(MailItemRef item)
        {
            if (item == null) return new AttachmentInfo[0];

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return new AttachmentInfo[0]; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return new AttachmentInfo[0];

                    var result = new List<AttachmentInfo>();
                    using (var attachments = new ComRef<OutlookOM.Attachments>(mail.Attachments))
                    {
                        int count = attachments.Value.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            OutlookOM.Attachment raw;
                            try { raw = attachments.Value[i]; } catch { continue; }
                            using (var att = new ComRef<OutlookOM.Attachment>(raw))
                            {
                                bool inline = IsInlineAttachment(att.Value);
                                string name = Safe(() => att.Value.FileName, string.Empty);
                                long size = Safe(() => (long)att.Value.Size, 0L);
                                int id = Safe(() => att.Value.Index, i);
                                result.Add(new AttachmentInfo(id, name, size, inline));
                            }
                        }
                    }
                    return result;
                }
            }
        }

        /// <summary>
        /// True when the attachment is inline/embedded (a cid:-linked body image,
        /// signature logo, or embedded/OLE item) rather than a true detached file
        /// (v2.5.0.0 B2). Any one signal is sufficient:
        /// 1. Type is olEmbeddeditem or olOLE (embedded objects, never standalone files).
        /// 2. PR_ATTACHMENT_HIDDEN (0x7FFE000B) true - Outlook's hidden-attachment flag.
        /// 3. A content-id (PR_ATTACH_CONTENT_ID) AND the ATT_MHTML_REF bit in
        ///    PR_ATTACH_FLAGS - i.e. actually referenced from the HTML body. A
        ///    content-id ALONE is not enough: Outlook assigns one to ordinary
        ///    attachments too, so checking it alone wrongly hides real files.
        /// olByValue (normal files) and olByReference (linked files) are NOT inline.
        /// </summary>
        private static bool IsInlineAttachment(OutlookOM.Attachment att)
        {
            var type = Safe(() => att.Type, OutlookOM.OlAttachmentType.olByValue);
            if (type == OutlookOM.OlAttachmentType.olEmbeddeditem
                || type == OutlookOM.OlAttachmentType.olOLE)
                return true;

            // The flags below need the PropertyAccessor, which not every store
            // exposes for every MAPI tag (PST vs Exchange differ), so failures
            // fall through to "not inline" (the attachment is shown, as pre-B2).
            try
            {
                using (var pa = new ComRef<OutlookOM.PropertyAccessor>(att.PropertyAccessor))
                {
                    const string PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";
                    if (Safe(() => (bool)pa.Value.GetProperty(PR_ATTACHMENT_HIDDEN), false))
                        return true;

                    const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
                    const string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";
                    const int ATT_MHTML_REF = 0x00000004;
                    string cid = Safe(() => pa.Value.GetProperty(PR_ATTACH_CONTENT_ID) as string, null);
                    int flags = Safe(() => Convert.ToInt32(pa.Value.GetProperty(PR_ATTACH_FLAGS)), 0);
                    if (!string.IsNullOrEmpty(cid) && (flags & ATT_MHTML_REF) != 0)
                        return true;
                }
            }
            catch { /* property not available for this store/item - treat as not inline */ }

            return false;
        }

        public bool SaveAttachmentToFile(MailItemRef item, int attachmentId, string destinationDirectory)
        {
            if (item == null || string.IsNullOrWhiteSpace(destinationDirectory)) return false;

            // Outlook's SaveAsFile throws an opaque "file not found" when the
            // directory is absent; create it (covers a OneDrive folder not yet
            // materialised locally) and bail clearly if that is not possible.
            try { System.IO.Directory.CreateDirectory(destinationDirectory); }
            catch (Exception ex)
            {
                Log.Error(ex, "SaveAttachmentToFile: destination directory unusable: {Dir}", destinationDirectory);
                return false;
            }

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return false; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return false;

                    using (var attachments = new ComRef<OutlookOM.Attachments>(mail.Attachments))
                    {
                        int count = attachments.Value.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            OutlookOM.Attachment raw;
                            try { raw = attachments.Value[i]; } catch { continue; }
                            using (var att = new ComRef<OutlookOM.Attachment>(raw))
                            {
                                if (Safe(() => att.Value.Index, -1) != attachmentId) continue;
                                string fileName = Safe(() => att.Value.FileName, null);
                                if (string.IsNullOrWhiteSpace(fileName))
                                    fileName = "attachment" + attachmentId;

                                string basePath = System.IO.Path.Combine(destinationDirectory, fileName);
                                if (!System.IO.File.Exists(basePath))
                                {
                                    try
                                    {
                                        att.Value.SaveAsFile(basePath);
                                        Log.Information("Saved attachment {Id} to {Path}.", attachmentId, basePath);
                                        return true;
                                    }
                                    catch (Exception ex)
                                    {
                                        Log.Error(ex, "SaveAttachmentToFile failed for attachment {Id} to {Path}.",
                                            attachmentId, basePath);
                                        return false;
                                    }
                                }

                                // A same-name file already exists: write to a same-directory temp
                                // first (same NTFS volume => the eventual rename is O(1)), then
                                // hash-compare rather than blindly suffixing (v2.9.0.0 D2).
                                string tempPath = System.IO.Path.Combine(destinationDirectory,
                                    "__rblclass_tmp_" + System.IO.Path.GetRandomFileName());
                                try
                                {
                                    att.Value.SaveAsFile(tempPath);
                                }
                                catch (Exception ex)
                                {
                                    try { if (System.IO.File.Exists(tempPath)) System.IO.File.Delete(tempPath); } catch { }
                                    Log.Error(ex, "SaveAttachmentToFile: temp-write failed for attachment {Id}.", attachmentId);
                                    return false;
                                }

                                bool identical;
                                try { identical = AttachmentChecksumHelper.ContentEquals(basePath, tempPath); }
                                catch (Exception ex)
                                {
                                    identical = false;
                                    Log.Warning(ex, "SaveAttachmentToFile: hash compare failed for {Path}; using suffix fallback.", basePath);
                                }

                                if (identical)
                                {
                                    try { System.IO.File.Delete(tempPath); } catch { }
                                    Log.Information(
                                        "SaveAttachmentToFile: dedupe skip - attachment {Id} already exists at {Path} with identical content.",
                                        attachmentId, basePath);
                                    return true;
                                }

                                string destPath = AttachmentFilenameCollisionResolver.Resolve(
                                    destinationDirectory, fileName, System.IO.File.Exists);
                                try
                                {
                                    System.IO.File.Move(tempPath, destPath);
                                    Log.Information("Saved attachment {Id} to {Path}.", attachmentId, destPath);
                                    return true;
                                }
                                catch (Exception ex)
                                {
                                    try { if (System.IO.File.Exists(tempPath)) System.IO.File.Delete(tempPath); } catch { }
                                    Log.Error(ex, "SaveAttachmentToFile: rename temp to {Path} failed for attachment {Id}.",
                                        destPath, attachmentId);
                                    return false;
                                }
                            }
                        }
                    }
                }
            }
            Log.Warning("SaveAttachmentToFile: attachment {Id} not found.", attachmentId);
            return false;
        }

        public bool SaveMailAsFile(MailItemRef item, string destinationDirectory)
        {
            if (item == null || string.IsNullOrWhiteSpace(destinationDirectory)) return false;

            try { System.IO.Directory.CreateDirectory(destinationDirectory); }
            catch (Exception ex)
            {
                Log.Error(ex, "SaveMailAsFile: destination directory unusable: {Dir}", destinationDirectory);
                return false;
            }

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return false; }

                using (var comItem = new ComRef<object>(rawItem))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return false;

                    string subject = Safe(() => mail.Subject, string.Empty) ?? string.Empty;
                    string baseName = SanitizeFileName(subject);
                    if (string.IsNullOrWhiteSpace(baseName)) baseName = "email";
                    string fileName = baseName + ".msg";

                    string basePath = System.IO.Path.Combine(destinationDirectory, fileName);
                    if (!System.IO.File.Exists(basePath))
                    {
                        try
                        {
                            mail.SaveAs(basePath, OutlookOM.OlSaveAsType.olMSG);
                            Log.Information("SaveMailAsFile: saved to {Path}", basePath);
                            return true;
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "SaveMailAsFile: SaveAs failed for {Path}", basePath);
                            return false;
                        }
                    }

                    // Same dedupe-aware save as SaveAttachmentToFile (v2.9.0.0 D2).
                    string tempPath = System.IO.Path.Combine(destinationDirectory,
                        "__rblclass_tmp_" + System.IO.Path.GetRandomFileName());
                    try
                    {
                        mail.SaveAs(tempPath, OutlookOM.OlSaveAsType.olMSG);
                    }
                    catch (Exception ex)
                    {
                        try { if (System.IO.File.Exists(tempPath)) System.IO.File.Delete(tempPath); } catch { }
                        Log.Error(ex, "SaveMailAsFile: temp-write failed for {Path}", tempPath);
                        return false;
                    }

                    bool identical;
                    try { identical = AttachmentChecksumHelper.ContentEquals(basePath, tempPath); }
                    catch (Exception ex)
                    {
                        identical = false;
                        Log.Warning(ex, "SaveMailAsFile: hash compare failed for {Path}; using suffix fallback.", basePath);
                    }

                    if (identical)
                    {
                        try { System.IO.File.Delete(tempPath); } catch { }
                        Log.Information(
                            "SaveMailAsFile: dedupe skip - {Path} already exists with identical content.", basePath);
                        return true;
                    }

                    string destPath = AttachmentFilenameCollisionResolver.Resolve(
                        destinationDirectory, fileName, System.IO.File.Exists);
                    try
                    {
                        System.IO.File.Move(tempPath, destPath);
                        Log.Information("SaveMailAsFile: saved to {Path}", destPath);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        try { if (System.IO.File.Exists(tempPath)) System.IO.File.Delete(tempPath); } catch { }
                        Log.Error(ex, "SaveMailAsFile: rename temp to {Path} failed.", destPath);
                        return false;
                    }
                }
            }
        }

        private static string SanitizeFileName(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject)) return string.Empty;
            char[] invalid = System.IO.Path.GetInvalidFileNameChars();
            var sb = new System.Text.StringBuilder();
            foreach (char c in subject)
            {
                if (Array.IndexOf(invalid, c) < 0) sb.Append(c);
            }
            string result = sb.ToString().Trim();
            // Truncate so directory + filename stays within 260-char MAX_PATH headroom.
            if (result.Length > 128) result = result.Substring(0, 128).TrimEnd();
            return result;
        }

        public bool StripExternalBanner(MailItemRef item, string bannerSignature)
        {
            if (item == null || string.IsNullOrWhiteSpace(bannerSignature)) return false;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                object rawItem;
                try { rawItem = session.Value.GetItemFromID(item.EntryId, item.StoreId); }
                catch { return false; }

                using (var comItem = new ComRef<object>(rawItem))
                    return StripBannerOnLiveMail(comItem.Value as OutlookOM.MailItem, bannerSignature);
            }
        }

        public bool StripBannerFromDraft(object draft, string bannerSignature)
        {
            if (string.IsNullOrWhiteSpace(bannerSignature)) return false;
            return StripBannerOnLiveMail(draft as OutlookOM.MailItem, bannerSignature);
        }

        /// <summary>
        /// Shared body-strip for a live <c>MailItem</c>: skip encrypted mail,
        /// run the Core stripper over the HTML body, write back and save only
        /// when it changed. Best-effort - logs and returns false on any fault.
        /// </summary>
        private static bool StripBannerOnLiveMail(OutlookOM.MailItem mail, string bannerSignature)
        {
            if (mail == null) return false;

            string entryId = Safe(() => mail.EntryID, null) ?? "(unsaved draft)";

            // Never rewrite an encrypted/signed body - it would corrupt it.
            string messageClass = Safe(() => mail.MessageClass, string.Empty);
            if (messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase))
            {
                Log.Information("StripExternalBanner skipped an encrypted/signed item ({MessageClass}, {EntryId}).",
                    messageClass, entryId);
                return false;
            }

            string html = Safe(() => mail.HTMLBody, null);
            if (string.IsNullOrEmpty(html)) return false;

            bool stripped;
            string updated = ExternalBannerStripper.Strip(html, bannerSignature, out stripped);
            if (!stripped) return false;

            try
            {
                mail.HTMLBody = updated;
                mail.Save();
                Log.Information("Stripped the external-sender banner from mail body {EntryId}.", entryId);
                return true;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Writing the banner-stripped body failed for {EntryId}.", entryId);
                return false;
            }
        }

        public string GetSelectedItemHtmlBody()
        {
            OutlookOM.Explorer rawExplorer;
            try { rawExplorer = _app.ActiveExplorer(); }
            catch { rawExplorer = null; }
            if (rawExplorer == null) return null;

            using (var explorer = new ComRef<OutlookOM.Explorer>(rawExplorer))
            {
                OutlookOM.Selection rawSelection;
                try { rawSelection = explorer.Value.Selection; }
                catch { return null; }

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
                            var mail = comItem.Value as OutlookOM.MailItem;
                            if (mail == null) continue; // skip non-mail
                            return Safe(() => mail.HTMLBody, null);
                        }
                    }
                }
            }

            return null;
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

        public MailItemRef ResolveMailItem(object item)
        {
            var mail = item as OutlookOM.MailItem;
            if (mail == null) return null;

            string entryId = Safe(() => mail.EntryID, null);
            if (entryId == null) return null;

            if (!TryGetFolderInfo(mail, out string storeId, out _)) return null;

            return new MailItemRef(storeId, entryId, Safe(() => mail.Subject, string.Empty));
        }

        public FolderNode GetFolderByEntryId(string storeId, string entryId)
        {
            if (string.IsNullOrEmpty(storeId) || string.IsNullOrEmpty(entryId)) return null;

            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                OutlookOM.Folder rawFolder;
                try { rawFolder = (OutlookOM.Folder)session.Value.GetFolderFromID(entryId, storeId); }
                catch { return null; } // no longer resolves - tolerate the miss

                using (var folder = new ComRef<OutlookOM.Folder>(rawFolder))
                {
                    // A folder "deleted" in Outlook is soft-moved under Deleted
                    // Items and still resolves by EntryID. For the callers of
                    // this method (the sent-item archive target) such a folder
                    // is gone - filing into it hides mail inside Deleted Items
                    // (v2.7.1.0 C1).
                    var rawStore = Safe<OutlookOM.Store>(() => folder.Value.Store, null);
                    if (rawStore != null)
                    {
                        using (var store = new ComRef<OutlookOM.Store>(rawStore))
                        {
                            string deletedItemsEntryId = SafeDeletedItemsEntryId(store.Value);
                            if (IsUnderDeletedItems(folder.Value, deletedItemsEntryId))
                                return null;
                        }
                    }
                    return DescribeFolder(folder.Value);
                }
            }
        }

        /// <summary>
        /// True when <paramref name="folder"/> is, or sits anywhere under, the
        /// store's Deleted Items folder - detected by walking the parent chain
        /// and comparing EntryIDs (robust against renames and localization).
        /// Does not release <paramref name="folder"/> itself (v2.7.1.0 C1).
        /// </summary>
        private static bool IsUnderDeletedItems(OutlookOM.Folder folder, string deletedItemsEntryId)
        {
            if (string.IsNullOrEmpty(deletedItemsEntryId)) return false;

            OutlookOM.Folder current = folder;   // caller-owned; never released here
            ComRef<OutlookOM.Folder> owned = null; // the parent we currently hold
            try
            {
                for (int depth = 0; depth < 64; depth++)
                {
                    string entryId = Safe(() => current.EntryID, null);
                    if (entryId == null) return false;
                    if (entryId == deletedItemsEntryId) return true;

                    object rawParent;
                    try { rawParent = current.Parent; }
                    catch { return false; }

                    var parentFolder = rawParent as OutlookOM.Folder;
                    if (parentFolder == null)
                    {
                        // Reached the store root (Parent is a Store/NameSpace).
                        if (rawParent != null) new ComRef<object>(rawParent).Dispose();
                        return false;
                    }

                    owned?.Dispose();
                    owned = new ComRef<OutlookOM.Folder>(parentFolder);
                    current = parentFolder;
                }
                return false;
            }
            finally { owned?.Dispose(); }
        }

        public FolderNode PickFolder()
        {
            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                OutlookOM.Folder rawFolder;
                try { rawFolder = (OutlookOM.Folder)session.Value.PickFolder(); }
                catch { return null; }

                if (rawFolder == null) return null; // user cancelled

                using (var folder = new ComRef<OutlookOM.Folder>(rawFolder))
                    return DescribeFolder(folder.Value);
            }
        }

        /// <summary>Build a FolderNode from a live folder (v2.7.0.0 B2 archive picker/resolve).</summary>
        private FolderNode DescribeFolder(OutlookOM.Folder folder)
        {
            string storeId = Safe(() => folder.StoreID, null);
            string entryId = Safe(() => folder.EntryID, null);
            if (storeId == null || entryId == null) return null;

            string name = Safe(() => folder.Name, string.Empty);
            string fullPath = Safe(() => folder.FolderPath, name);
            int subCount = SafeSubfolderCount(folder);
            return new FolderNode(storeId, entryId, parentEntryId: null,
                                  name: name, fullPath: fullPath ?? name, isLeaf: subCount == 0);
        }

        public FolderNode GetInboxFolder()
        {
            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                ComRef<OutlookOM.Folder> inbox = null;
                try
                {
                    inbox = new ComRef<OutlookOM.Folder>(
                        (OutlookOM.Folder)session.Value.GetDefaultFolder(OutlookOM.OlDefaultFolders.olFolderInbox));

                    string storeId = Safe(() => inbox.Value.StoreID, null);
                    string entryId = Safe(() => inbox.Value.EntryID, null);
                    if (storeId == null || entryId == null) return null;

                    string name = Safe(() => inbox.Value.Name, "Inbox");
                    int subCount = SafeSubfolderCount(inbox.Value);
                    return new FolderNode(storeId, entryId, parentEntryId: null,
                                          name: name, fullPath: name, isLeaf: subCount == 0);
                }
                catch { return null; } // no default store / folder unavailable - tolerate
                finally { inbox?.Dispose(); }
            }
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

        /// <summary>
        /// Mark a resolved Outlook item (MailItem or MeetingItem) as read in-place.
        /// For MailItem, first checks MessageClass: S/MIME-encrypted/signed items
        /// (IPM.Note.SMIME*) are skipped because Save() on such items risks payload
        /// corruption — ClassifyRequest.MarkAsRead promises "S/MIME-safe".
        /// MeetingItem objects carry no S/MIME payload (IPM.Schedule.Meeting.*) and
        /// are marked unconditionally. Best-effort: all COM faults are swallowed
        /// (v2.7.1.0 A1).
        /// </summary>
        private static void TryMarkItemRead(OutlookOM.MailItem mail, OutlookOM.MeetingItem meeting)
        {
            if (mail != null)
            {
                string messageClass = Safe(() => mail.MessageClass, string.Empty);
                if (messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase))
                {
                    Log.Information(
                        "MarkAsRead skipped an encrypted/signed item ({MessageClass}).",
                        messageClass);
                    return;
                }
                try { mail.UnRead = false; mail.Save(); }
                catch (Exception ex) { Log.Warning(ex, "MarkAsRead failed for a filed mail."); }
            }
            else if (meeting != null)
            {
                try { meeting.UnRead = false; meeting.Save(); }
                catch (Exception ex) { Log.Warning(ex, "MarkAsRead failed for a filed meeting item."); }
            }
        }
    }
}
