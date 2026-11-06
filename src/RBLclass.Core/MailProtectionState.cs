namespace RBLclass.Core
{
    /// <summary>
    /// What the diagnostic trace records about a mail's protection (v2.11.0.0 A3).
    /// Outlook-free so <see cref="ClassifierService"/> can log it.
    /// </summary>
    public sealed class MailProtectionState
    {
        public MailProtectionState(string messageClass, int? securityFlags)
        {
            MessageClass = messageClass;
            SecurityFlags = securityFlags;
        }

        /// <summary>The message class, or null when it could not be read.</summary>
        public string MessageClass { get; }

        /// <summary>
        /// <c>PR_SECURITY_FLAGS</c>, or null when absent (the normal case on plain
        /// mail). Trace only: never used to decide protection.
        /// </summary>
        public int? SecurityFlags { get; }

        public bool IsProtected => ProtectedMailPolicy.IsProtected(MessageClass);
    }
}
