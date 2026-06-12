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
                               bool safetyCopy = false)
        {
            Items = (items ?? throw new ArgumentNullException(nameof(items))).ToArray();
            Destinations = (destinations ?? throw new ArgumentNullException(nameof(destinations))).ToArray();
            KeepCopy = keepCopy;
            RemoveAttachments = removeAttachments;
            MarkTasksComplete = markTasksComplete;
            SafetyCopy = safetyCopy;
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
    }
}
