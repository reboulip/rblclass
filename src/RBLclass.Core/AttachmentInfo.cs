namespace RBLclass.Core
{
    /// <summary>
    /// Outlook-free descriptor of one mail attachment (v2.4.0.0 F2): enough to
    /// list it in the disposition modal and to identify it for save/strip.
    /// </summary>
    public sealed class AttachmentInfo
    {
        public AttachmentInfo(int id, string fileName, long sizeBytes)
        {
            Id = id;
            FileName = fileName ?? string.Empty;
            SizeBytes = sizeBytes;
        }

        /// <summary>Position-based id (Outlook <c>Attachment.Index</c>), stable until attachments are removed.</summary>
        public int Id { get; }

        public string FileName { get; }

        public long SizeBytes { get; }
    }
}
