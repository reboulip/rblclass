namespace RBLclass.Core
{
    /// <summary>
    /// Result of a standalone attachment-disposition apply
    /// (v2.8.0.0 A2: <see cref="IClassifier.ApplyAttachmentDispositions"/>).
    /// </summary>
    public sealed class AttachmentsAppliedResult
    {
        public AttachmentsAppliedResult(int applied, int errors, int encryptedSkips)
        {
            Applied = applied;
            Errors = errors;
            EncryptedSkips = encryptedSkips;
        }

        /// <summary>Item-groups where at least one attachment was removed.</summary>
        public int Applied { get; }

        /// <summary>Item-groups skipped because a requested save failed (strip was blocked).</summary>
        public int Errors { get; }

        /// <summary>Item-groups where <see cref="IMailStore.RemoveAttachments"/> returned false (S/MIME guard).</summary>
        public int EncryptedSkips { get; }
    }
}
