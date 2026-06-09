using System;
using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Result of a folder search: the (possibly truncated) hit list plus the
    /// information the UI needs to react to an over-cap search.
    /// </summary>
    public sealed class FolderSearchOutcome
    {
        public static readonly FolderSearchOutcome Empty =
            new FolderSearchOutcome(new FolderSearchResult[0], 0, false);

        public FolderSearchOutcome(IReadOnlyList<FolderSearchResult> results,
                                   int totalMatchCount,
                                   bool limitExceeded)
        {
            Results = results ?? throw new ArgumentNullException(nameof(results));
            TotalMatchCount = totalMatchCount;
            LimitExceeded = limitExceeded;
        }

        /// <summary>The results to display, truncated to <see cref="FolderSearchOptions.MaxResults"/>.</summary>
        public IReadOnlyList<FolderSearchResult> Results { get; }

        /// <summary>Total number of hits before truncation.</summary>
        public int TotalMatchCount { get; }

        /// <summary>True when <see cref="TotalMatchCount"/> exceeded the cap and results were truncated.</summary>
        public bool LimitExceeded { get; }
    }
}
