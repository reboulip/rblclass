using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// The cached folder tree and its index lifecycle. The methods are split by
    /// the thread they must run on so the add-in shell can keep COM access on
    /// the Outlook UI thread while doing SQLite I/O off it:
    /// <list type="bullet">
    ///   <item><see cref="Load"/> - SQLite only, safe on a background thread.</item>
    ///   <item><see cref="WalkAndPersist"/> / <see cref="ReindexStore"/> - touch
    ///   the live store via <see cref="IMailStore"/>, so must run on the Outlook
    ///   UI thread.</item>
    /// </list>
    /// </summary>
    public interface IFolderTree
    {
        /// <summary>
        /// Ensure the schema exists and, if the index is already populated, load
        /// it into the in-memory cache. Returns <see cref="IndexSource.LoadedFromCache"/>
        /// when loaded, or <see cref="IndexSource.NeedsWalk"/> when the index is
        /// empty and a first-run walk is required. No Outlook access.
        /// </summary>
        IndexResult Load();

        /// <summary>
        /// Walk every non-excluded live store via <see cref="IMailStore"/>,
        /// persist the result, and populate the cache. Used on first run (and to
        /// force a full rebuild). Touches COM - call on the Outlook UI thread.
        /// </summary>
        IndexResult WalkAndPersist();

        /// <summary>
        /// Re-walk a single store and replace just its folders in the index and
        /// cache (targeted re-index after a sub-folder change). Touches COM.
        /// </summary>
        void ReindexStore(string storeId);

        /// <summary>A snapshot of the currently cached folders.</summary>
        IReadOnlyList<FolderNode> GetAll();

        /// <summary>
        /// Persist pre-walked store data and populate the in-memory cache in one
        /// transaction. Called from the shell's per-store async loop after all COM
        /// access is done; Core does no further COM calls. Does NOT transition
        /// IndexStatus — the caller owns Indexing → Ready.
        /// </summary>
        IndexResult PersistWalkedStores(
            IReadOnlyList<(StoreInfo Store, IReadOnlyList<FolderNode> Folders)> walkedStores);
    }
}
