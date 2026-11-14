namespace RBLclass.Core
{
    /// <summary>
    /// One attachment of a decrypted mail, with its bytes, as read by the adapter
    /// for re-encrypted filing (v2.12.0.0 A2). Outlook-free.
    /// </summary>
    public sealed class MailPart
    {
        /// <param name="attachmentId">
        /// The Outlook <c>Attachment.Index</c>: the same id <see cref="AttachmentInfo.Id"/>
        /// and <see cref="AttachmentDisposition.AttachmentId"/> carry, so a
        /// disposition row maps to its part explicitly, never by position.
        /// </param>
        /// <param name="data">
        /// Not copied (a large mail must not be held twice): callers treat it as
        /// immutable once passed in.
        /// </param>
        public MailPart(int attachmentId, string fileName, string contentType, byte[] data,
                        bool isInline = false, string contentId = null)
        {
            AttachmentId = attachmentId;
            FileName = fileName ?? string.Empty;
            ContentType = contentType;
            Data = data ?? new byte[0];
            IsInline = isInline;
            ContentId = isInline ? InlineImageRules.NormalizeContentId(contentId) : null;
        }

        public int AttachmentId { get; }

        public string FileName { get; }

        /// <summary>The MIME type (<c>PR_ATTACH_MIME_TAG</c>), or null when unknown.</summary>
        public string ContentType { get; }

        public byte[] Data { get; }

        /// <summary>
        /// Inline by the adapter's rule (body image, hidden part): never removed by
        /// "remove all attachments". May be true with <see cref="ContentId"/> null.
        /// </summary>
        public bool IsInline { get; }

        /// <summary>The normalised content id (no angle brackets) of an inline part, or null.</summary>
        public string ContentId { get; }
    }
}
