using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace RBLclass.Core
{
    /// <summary>
    /// The flagship action (legacy 5b): file mail items into one or more
    /// folders. All decision logic lives here and is unit-tested with a faked
    /// <see cref="IMailStore"/>; the Outlook mechanics live in the adapter.
    /// </summary>
    public interface IClassifier
    {
        /// <summary>
        /// Compute the final item set and surface anything the caller should
        /// confirm before filing (legacy 5b steps 2-3): optionally widens
        /// <paramref name="items"/> with conversation siblings, dedupes by
        /// (StoreId, EntryId), and reports which of the resulting items are
        /// flagged-incomplete tasks. Call this before building a
        /// <see cref="ClassifyRequest"/> so the UI can prompt when
        /// <see cref="ClassifyPreflight.FlaggedIncomplete"/> is non-empty.
        /// </summary>
        ClassifyPreflight Preflight(IReadOnlyList<MailItemRef> items, bool widenConversation);

        ClassifyResult Classify(ClassifyRequest request);

        /// <summary>
        /// Responsive variant of <see cref="Classify"/> (v2.4 D2): same filing
        /// semantics, undo plan and item order, but reports progress once per
        /// processed item via <paramref name="progress"/> and awaits
        /// <paramref name="yieldBetweenItems"/> between items so the caller can
        /// keep its (STA) message pump alive - Outlook repaints and stays
        /// interactive, and an inline mail scanner (Stormshield) gets a stable
        /// window between moves (closes D1). The non-COM history write is
        /// dispatched off the caller's thread. Pass a no-op yield in tests.
        /// </summary>
        Task<ClassifyResult> ClassifyAsync(ClassifyRequest request,
                                           IProgress<ClassifyProgress> progress,
                                           Func<Task> yieldBetweenItems);

        /// <summary>
        /// Reverse a previous classify (v2.2 Undo): re-mark un-completed flags,
        /// permanently delete the copies it created, and move the moved items
        /// back to their source folders. Best-effort per action - one failure
        /// is counted and the rest still execute. Stripped attachments cannot
        /// be restored (<see cref="ClassifyUndoPlan.AttachmentStrips"/>).
        /// </summary>
        UndoResult Undo(ClassifyUndoPlan plan);

        /// <summary>
        /// Auto-class (v2.2): file each of <paramref name="items"/> to its
        /// conversation's most recently recorded destination(s), looked up in
        /// the classification history and validated by
        /// <paramref name="resolveLiveFolder"/> (which returns the live
        /// <see cref="FolderNode"/> for a (storeId, entryId), or null when that
        /// folder no longer exists in the index). Items with no history, or
        /// whose every remembered destination is stale, are skipped and counted.
        /// Honours the same <paramref name="keepCopy"/>/<paramref name="removeAttachments"/>/<paramref name="safetyCopy"/>
        /// options as a manual classify and produces one combined undo plan.
        /// Requires that this service was constructed with a history store.
        /// </summary>
        /// <param name="attachmentDispositions">
        /// Per-attachment choices collected by the caller's modal pre-flight.
        /// When non-null, the classifier honours each per-attachment action
        /// (matching the Modal-mode manual classify flow). When null, all
        /// attachments are stripped silently (DeleteSilently behaviour).
        /// </param>
        AutoClassifyResult AutoClassify(IReadOnlyList<MailItemRef> items,
                                        Func<string, string, FolderNode> resolveLiveFolder,
                                        bool keepCopy, bool removeAttachments, bool safetyCopy,
                                        IReadOnlyList<AttachmentDisposition> attachmentDispositions = null,
                                        AttachmentLabelOptions labelOptions = null);
    }
}
