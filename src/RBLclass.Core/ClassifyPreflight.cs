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
                                 IReadOnlyList<MailItemRef> flaggedIncomplete,
                                 IReadOnlyList<string> skippedEncrypted = null,
                                 bool wideningTableDisabledThisSession = false)
        {
            Items = items;
            FlaggedIncomplete = flaggedIncomplete;
            SkippedEncrypted = skippedEncrypted ?? new string[0];
            WideningTableDisabledThisSession = wideningTableDisabledThisSession;
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

        /// <summary>
        /// Subjects of conversation siblings that were skipped because they are
        /// encrypted/signed (the encryption provider is inactive). Empty unless
        /// widening was requested and an in-scope encrypted sibling was found -
        /// the caller should warn the user these were left in place.
        /// </summary>
        public IReadOnlyList<string> SkippedEncrypted { get; }

        /// <summary>
        /// True on the first preflight in a session where the Outlook
        /// conversation table timed out and the store switched to its
        /// folder-scan fallback - the caller should surface a one-time notice
        /// (B3, v2.9.0.0).
        /// </summary>
        public bool WideningTableDisabledThisSession { get; }
    }
}
