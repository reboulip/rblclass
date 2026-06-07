using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using RBLclass.Core;
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
