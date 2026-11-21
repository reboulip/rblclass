namespace RBLclass.Core
{
    /// <summary>
    /// What the attachment warning shows (v2.13.0.0 D2): the sender and, when the
    /// event identifies it, the attachment file name. (BeforeAttachmentPreview carries
    /// no attachment, so the name can be empty.)
    /// </summary>
    public sealed class PhishingWarningMessage
    {
        public PhishingWarningMessage(SenderInfo sender, string attachmentFileName)
        {
            Sender = sender ?? SenderInfo.Unknown;
            AttachmentFileName = attachmentFileName ?? string.Empty;
        }

        public SenderInfo Sender { get; }
        public string AttachmentFileName { get; }
        public bool HasAttachmentFileName => AttachmentFileName.Length > 0;
    }
}
