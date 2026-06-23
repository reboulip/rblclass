namespace RBLclass.Core
{
    /// <summary>
    /// Outlook-free descriptor of one mail attachment (v2.4.0.0 F2): enough to
    /// list it in the disposition modal and to identify it for save/strip.
    /// </summary>
    public sealed class AttachmentInfo
    {
        public AttachmentInfo(int id, string fileName, long sizeBytes, bool isInline = false)
        {
            Id = id;
            FileName = fileName ?? string.Empty;
            SizeBytes = sizeBytes;
            IsInline = isInline;
        }

        /// <summary>Position-based id (Outlook <c>Attachment.Index</c>), stable until attachments are removed.</summary>
        public int Id { get; }

        public string FileName { get; }

        public long SizeBytes { get; }

        /// <summary>
        /// True when this is an inline/embedded attachment (a cid:-linked body
        /// image, signature logo, or OLE/embedded item) rather than a true
        /// detached file (v2.5.0.0 B2). Such attachments are not offered for
        /// disposition and are never stripped. Populated by the Outlook adapter.
        /// </summary>
        public bool IsInline { get; }
    }
}
