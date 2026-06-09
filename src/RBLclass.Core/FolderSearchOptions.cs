namespace RBLclass.Core
{
    /// <summary>
    /// Options for a folder search. Mirrors the legacy toggles (All results,
    /// MaxResults) plus the new <see cref="FolderMatchMode"/>. These map onto
    /// user preferences in Step 9.
    /// </summary>
    public sealed class FolderSearchOptions
    {
        /// <summary>Legacy default cap on displayed results.</summary>
        public const int DefaultMaxResults = 100;

        public FolderSearchOptions(
            FolderMatchMode matchMode = FolderMatchMode.WordPrefix,
            bool allResults = false,
            int maxResults = DefaultMaxResults)
        {
            MatchMode = matchMode;
            AllResults = allResults;
            MaxResults = maxResults < 1 ? 1 : maxResults;
        }

        /// <summary>How keywords are tested against folder paths.</summary>
        public FolderMatchMode MatchMode { get; }

        /// <summary>
        /// When false (default), a matched non-leaf folder is collapsed to a
        /// single "path | [...]" entry standing for its whole subtree. When
        /// true, every matching folder is listed individually.
        /// </summary>
        public bool AllResults { get; }

        /// <summary>
        /// Cap on the number of results returned. When the total exceeds this,
        /// the outcome is flagged (<see cref="FolderSearchOutcome.LimitExceeded"/>)
        /// and the list is truncated - the legacy tool refused to display and
        /// asked the user to refine.
        /// </summary>
        public int MaxResults { get; }

        public static FolderSearchOptions Default { get; } = new FolderSearchOptions();
    }
}
