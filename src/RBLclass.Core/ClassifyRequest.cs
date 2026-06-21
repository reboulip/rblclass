using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// A request to classify (file) a set of mail items into one or more
    /// destination folders. Mirrors the legacy classify toggles.
    /// </summary>
    /// <remarks>
    /// <see cref="Items"/> is expected to already be the final, widened set -
    /// the caller runs <see cref="IClassifier.Preflight"/> first, which both
    /// widens by conversation and reports which items need the task-completion
    /// prompt before the request is built. <c>Classify</c> itself does not widen.
    /// </remarks>
    public sealed class ClassifyRequest
    {
        public ClassifyRequest(IReadOnlyList<MailItemRef> items,
                               IReadOnlyList<FolderNode> destinations,
                               bool keepCopy,
                               bool removeAttachments,
                               bool markTasksComplete = false,
                               bool safetyCopy = false,
                               string bannerSignature = null,
                               IReadOnlyList<AttachmentDisposition> attachmentDispositions = null,
                               AttachmentLabelOptions labelOptions = null)
        {
            Items = (items ?? throw new ArgumentNullException(nameof(items))).ToArray();
            Destinations = (destinations ?? throw new ArgumentNullException(nameof(destinations))).ToArray();
            KeepCopy = keepCopy;
            RemoveAttachments = removeAttachments;
            MarkTasksComplete = markTasksComplete;
            SafetyCopy = safetyCopy;
            BannerSignature = bannerSignature;
            AttachmentDispositions = attachmentDispositions != null
                ? attachmentDispositions.ToArray()
                : (IReadOnlyList<AttachmentDisposition>)new AttachmentDisposition[0];
            LabelOptions = labelOptions;
        }

        /// <summary>The mail items to file (already widened, if requested).</summary>
        public IReadOnlyList<MailItemRef> Items { get; }

        /// <summary>
        /// Destination folders (1..N). Filing to N folders produces N copies per
        /// item (legacy copy-per-destination semantics).
        /// </summary>
        public IReadOnlyList<FolderNode> Destinations { get; }

        /// <summary>When true, originals are left in place (legacy "keep a copy").</summary>
        public bool KeepCopy { get; }

        /// <summary>When true, attachments are stripped from each item before filing.</summary>
        public bool RemoveAttachments { get; }

        /// <summary>
        /// When true, the filed copy of each flagged-incomplete item is marked
        /// task-complete (legacy "task-completion guard", confirmed by the
        /// caller via <see cref="ClassifyPreflight.FlaggedIncomplete"/> before
        /// building this request). Acts on the filed copy only, mirroring
        /// attachment removal - the original is left untouched.
        /// </summary>
        public bool MarkTasksComplete { get; }

        /// <summary>
        /// When true (and <see cref="KeepCopy"/> is off), each successfully
        /// moved original also leaves a copy in its source store's Deleted
        /// Items - the old delete-after-copy side effect, restored as an
        /// opt-in guardrail (v2.2 setting). The copy is taken from the moved
        /// item at its destination (out of the displayed folder, so it never
        /// re-creates the transient-item race other add-ins choked on) and
        /// BEFORE attachment stripping, so the guardrail copy keeps its
        /// attachments. A safety-copy failure never fails the classify.
        /// </summary>
        public bool SafetyCopy { get; }

        /// <summary>
        /// When non-empty, the learned external-sender banner is stripped from
        /// each filed item's HTML body (v2.2). Filed-copy-only, like attachment
        /// removal: with "keep a copy" on the kept original is untouched; with
        /// it off the moved item is edited at its destination. Encrypted items
        /// are never edited. The body change is not reversible by Undo.
        /// </summary>
        public string BannerSignature { get; }

        /// <summary>True when a banner strip is requested (a signature is present).</summary>
        public bool StripBanner => !string.IsNullOrWhiteSpace(BannerSignature);

        /// <summary>
        /// Per-attachment choices from the F2 disposition modal (v2.4.0.0), keyed
        /// to the original items by (StoreId, EntryId). Empty when the modal was
        /// not used (DeleteSilently mode, or no attachments): the classifier then
        /// falls back to stripping all attachments. Acts on the filed copy /
        /// moved item, mirroring the other attachment rules.
        /// </summary>
        public IReadOnlyList<AttachmentDisposition> AttachmentDispositions { get; }

        /// <summary>
        /// Localized templates for the F3 "former attachments" label written on
        /// each filed copy after a disposition (v2.4.0.0). Null = no label.
        /// </summary>
        public AttachmentLabelOptions LabelOptions { get; }
    }
}
