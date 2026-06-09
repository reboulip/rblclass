namespace RBLclass.Core
{
    /// <summary>Where the current folder index came from.</summary>
    public enum IndexSource
    {
        /// <summary>No persisted index yet - a live walk is needed.</summary>
        NeedsWalk = 0,

        /// <summary>Loaded from the SQLite cache (no Outlook walk).</summary>
        LoadedFromCache = 1,

        /// <summary>Freshly walked from the live Outlook stores and persisted.</summary>
        Walked = 2
    }

    /// <summary>Outcome of an index lifecycle operation, for logging/diagnostics.</summary>
    public sealed class IndexResult
    {
        public IndexResult(IndexSource source, int storeCount, int folderCount)
        {
            Source = source;
            StoreCount = storeCount;
            FolderCount = folderCount;
        }

        public IndexSource Source { get; }

        /// <summary>Number of stores represented in the index.</summary>
        public int StoreCount { get; }

        /// <summary>Number of folders in the index.</summary>
        public int FolderCount { get; }
    }
}
