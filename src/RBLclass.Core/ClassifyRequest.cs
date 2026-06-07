using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// A request to classify (file) a set of mail items into one or more
    /// destination folders. Mirrors the legacy classify toggles. Conversation
    /// widening and the task-completion guard are deferred to Step 6 and are not
    /// part of this request yet.
    /// </summary>
    public sealed class ClassifyRequest
    {
        public ClassifyRequest(IReadOnlyList<MailItemRef> items,
                               IReadOnlyList<FolderNode> destinations,
                               bool keepCopy,
                               bool removeAttachments)
        {
            Items = (items ?? throw new ArgumentNullException(nameof(items))).ToArray();
            Destinations = (destinations ?? throw new ArgumentNullException(nameof(destinations))).ToArray();
            KeepCopy = keepCopy;
            RemoveAttachments = removeAttachments;
        }

        /// <summary>The mail items to file.</summary>
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
    }
}
