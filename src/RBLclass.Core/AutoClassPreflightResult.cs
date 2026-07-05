using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Result of an Auto-class history preflight (v2.8.0.0 C1): which selected
    /// items can actually be filed, and per-reason counts for those that cannot.
    /// Lets the pane skip the attachment-disposition modal for items with no
    /// live remembered destination.
    /// </summary>
    public sealed class AutoClassPreflightResult
    {
        public AutoClassPreflightResult(
            IReadOnlyList<MailItemRef> processable,
            int noHistory, int staleFolders, int alreadyInPlace)
        {
            Processable = (processable ?? new MailItemRef[0]).ToArray();
            NoHistory = noHistory;
            StaleFolders = staleFolders;
            AlreadyInPlace = alreadyInPlace;
        }

        /// <summary>Items with a valid, live remembered destination they are not already in.</summary>
        public IReadOnlyList<MailItemRef> Processable { get; }

        /// <summary>Items skipped because the conversation was never classified (or history fell outside the retention window).</summary>
        public int NoHistory { get; }

        /// <summary>Items skipped because every remembered destination folder is no longer in the live index.</summary>
        public int StaleFolders { get; }

        /// <summary>Items skipped because the item is already in its only remembered destination folder.</summary>
        public int AlreadyInPlace { get; }

        public bool HasProcessableItems => Processable.Count > 0;
    }
}
