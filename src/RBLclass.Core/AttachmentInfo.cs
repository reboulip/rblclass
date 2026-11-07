namespace RBLclass.Core
{
    /// <summary>
    /// Outlook-free descriptor of one mail attachment (v2.4.0.0 F2): enough to
    /// list it in the disposition modal and to identify it for save/strip.
    /// </summary>
    public sealed class AttachmentInfo
    {
        public AttachmentInfo(int id, string fileName, long sizeBytes, bool isInline = false, string contentId = null)
        {
            Id = id;
            FileName = fileName ?? string.Empty;
            SizeBytes = sizeBytes;
            IsInline = isInline;
            ContentId = isInline ? InlineImageRules.NormalizeContentId(contentId) : null;
        }

        /// <summary>Position-based id (Outlook <c>Attachment.Index</c>), stable until attachments are removed.</summary>
        public int Id { get; }

        public string FileName { get; }

        public long SizeBytes { get; }

        /// <summary>
        /// True when this is an inline/embedded attachment (a cid:-linked body
        /// image, signature logo, or OLE/embedded item) rather than a true
        /// detached file (v2.5.0.0 B2). Only a body image with a content id can be
        /// offered for disposition, and only when the user opted in (v2.11.0.0
        /// C2); none is ever stripped by "remove all attachments". Populated by
        /// the Outlook adapter.
        /// </summary>
        public bool IsInline { get; }

        /// <summary>The content id a body image is referenced by (<c>cid:</c>), or null.</summary>
        public string ContentId { get; }

        /// <summary>A body image that can be removed by rewriting its reference (v2.11.0.0 C2).</summary>
        public bool IsInlineImage => IsInline && ContentId != null;
    }
}
