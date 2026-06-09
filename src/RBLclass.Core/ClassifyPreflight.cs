using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Result of <see cref="IClassifier.Preflight"/>: the (possibly widened) set
    /// of items a classify will act on, plus the subset flagged as an incomplete
    /// task - so the caller can prompt before committing to the move (legacy
    /// 5b steps 2-3, now decoupled from the actual filing).
    /// </summary>
    public sealed class ClassifyPreflight
    {
        public ClassifyPreflight(IReadOnlyList<MailItemRef> items,
                                 IReadOnlyList<MailItemRef> flaggedIncomplete)
        {
            Items = items;
            FlaggedIncomplete = flaggedIncomplete;
        }

        /// <summary>
        /// The items the classify will file - the original selection, widened
        /// with conversation siblings and deduped by (StoreId, EntryId) when
        /// requested.
        /// </summary>
        public IReadOnlyList<MailItemRef> Items { get; }

        /// <summary>
        /// The subset of <see cref="Items"/> flagged as a not-yet-completed task
        /// (legacy "task-completion guard"). Empty when none qualify - the
        /// caller should only prompt when this is non-empty.
        /// </summary>
        public IReadOnlyList<MailItemRef> FlaggedIncomplete { get; }
    }
}
