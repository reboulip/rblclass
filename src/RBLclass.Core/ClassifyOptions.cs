namespace RBLclass.Core
{
    /// <summary>
    /// Immutable set of per-action toggle flags for a classify operation.
    /// Replaces the positional bool parameters that were growing the
    /// <see cref="ClassifyRequest"/> constructor each sprint (v2.7.1.0 F1).
    /// </summary>
    public sealed class ClassifyOptions
    {
        /// <summary>All toggles at their defaults: move (no keep), no strip, no task-complete, no safety copy, not marked read, record history.</summary>
        public static readonly ClassifyOptions Default = new ClassifyOptions();

        public ClassifyOptions(
            bool keepCopy = false,
            bool removeAttachments = false,
            bool markTasksComplete = false,
            bool safetyCopy = false,
            bool markAsRead = false,
            bool recordHistory = true,
            bool reencryptProtected = false,
            string reencryptCertificateThumbprint = null)
        {
            KeepCopy = keepCopy;
            RemoveAttachments = removeAttachments;
            MarkTasksComplete = markTasksComplete;
            SafetyCopy = safetyCopy;
            MarkAsRead = markAsRead;
            RecordHistory = recordHistory;
            ReencryptProtected = reencryptProtected;
            ReencryptCertificateThumbprint = reencryptCertificateThumbprint ?? string.Empty;
        }

        public bool KeepCopy { get; }
        public bool RemoveAttachments { get; }
        public bool MarkTasksComplete { get; }
        public bool SafetyCopy { get; }
        public bool MarkAsRead { get; }
        public bool RecordHistory { get; }

        /// <summary>
        /// Experimental (v2.12.0.0): file encrypted mail readable decrypted in
        /// memory as a rebuilt, re-encrypted copy instead of moving it.
        /// </summary>
        public bool ReencryptProtected { get; }

        /// <summary>The certificate to re-encrypt to; empty = Automatic.</summary>
        public string ReencryptCertificateThumbprint { get; }
    }
}
