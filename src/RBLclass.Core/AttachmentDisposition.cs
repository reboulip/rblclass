using System;

namespace RBLclass.Core
{
    /// <summary>What to do with one attachment at classify time (v2.4.0.0 F2).</summary>
    public enum AttachmentDispositionAction
    {
        Delete = 0,
        SaveTo = 1,
        /// <summary>Leave the attachment on the filed copy untouched (v2.5.0.0 B1).</summary>
        Keep = 2
    }

    /// <summary>
    /// The user's choice for one attachment in the disposition modal (v2.4.0.0
    /// F2): delete it, or save it into a directory (then strip it from the filed
    /// mail). Keyed to the original item by (StoreId, EntryId); the classifier
    /// re-targets the filed copy / moved item. Carries <see cref="FileName"/> so
    /// the F3 label can describe it after the attachment is gone.
    /// </summary>
    public sealed class AttachmentDisposition
    {
        public AttachmentDisposition(MailItemRef item, int attachmentId, string fileName,
                                     AttachmentDispositionAction action, string targetDirectory = null,
                                     string contentId = null, string inlinePlaceholder = null)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
            AttachmentId = attachmentId;
            FileName = fileName ?? string.Empty;
            Action = action;
            TargetDirectory = targetDirectory ?? string.Empty;
            ContentId = InlineImageRules.NormalizeContentId(contentId);
            InlinePlaceholder = inlinePlaceholder ?? string.Empty;
        }

        /// <summary>Set for a body image (v2.11.0.0 C2): removed by rewriting the body, not by index.</summary>
        public string ContentId { get; }

        public bool IsInlineImage => ContentId != null;

        /// <summary>Text that replaces a removed body image.</summary>
        public string InlinePlaceholder { get; }

        public MailItemRef Item { get; }
        public int AttachmentId { get; }
        public string FileName { get; }
        public AttachmentDispositionAction Action { get; }

        /// <summary>Destination directory for a <see cref="AttachmentDispositionAction.SaveTo"/>; empty for Delete.</summary>
        public string TargetDirectory { get; }
    }
}
