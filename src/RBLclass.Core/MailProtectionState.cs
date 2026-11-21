namespace RBLclass.Core
{
    /// <summary>
    /// What the diagnostic trace records about a mail's protection (v2.11.0.0 A3).
    /// Outlook-free so <see cref="ClassifierService"/> can log it.
    /// </summary>
    public sealed class MailProtectionState
    {
        public MailProtectionState(string messageClass, int? securityFlags,
                                   string storedMessageClass = null,
                                   bool hasSmimePayload = false,
                                   bool hasUnsavedChanges = false)
        {
            MessageClass = messageClass;
            SecurityFlags = securityFlags;
            StoredMessageClass = storedMessageClass;
            HasSmimePayload = hasSmimePayload;
            HasUnsavedChanges = hasUnsavedChanges;
            Reason = ProtectedMailPolicy.Evaluate(new ProtectionSignals(
                messageClass, storedMessageClass, hasSmimePayload, hasUnsavedChanges));
        }

        /// <summary>The message class, or null when it could not be read.</summary>
        public string MessageClass { get; }

        /// <summary>
        /// <c>PR_SECURITY_FLAGS</c>, or null when absent (the normal case on plain
        /// mail). Trace only: never used to decide protection.
        /// </summary>
        public int? SecurityFlags { get; }

        /// <summary>The class stored in the folder, or null when not looked up / not found.</summary>
        public string StoredMessageClass { get; }

        public bool HasSmimePayload { get; }

        public bool HasUnsavedChanges { get; }

        /// <summary>The first signal that made the mail protected, or <see cref="ProtectionReason.None"/>.</summary>
        public ProtectionReason Reason { get; }

        public bool IsProtected => Reason != ProtectionReason.None;
    }
}
