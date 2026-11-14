using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Why a mail counts as protected (encrypted or signed), or <see cref="None"/>.
    /// </summary>
    public enum ProtectionReason
    {
        /// <summary>No signal says the mail is protected.</summary>
        None = 0,

        /// <summary>The message class could not be read: counted as protected.</summary>
        UnreadableClass,

        /// <summary>The live message class is an S/MIME or secure class.</summary>
        MessageClass,

        /// <summary>
        /// The class stored in the folder is an S/MIME or secure class while the
        /// opened item reports another one: an add-in (Stormshield with the smart
        /// card inserted) decrypted the item in memory (#62).
        /// </summary>
        StoredMessageClass,

        /// <summary>The item still carries its S/MIME payload part (smime.p7m / smime.p7s).</summary>
        SmimePayload,

        /// <summary>
        /// The item had unsaved changes as soon as it was opened, before RBLclass
        /// touched it: something rewrote it in memory, and saving would commit that
        /// rewrite (#62).
        /// </summary>
        UnsavedChanges,
    }

    /// <summary>
    /// What the adapter could read about a mail's protection. Outlook-free so the
    /// decision stays in Core and is unit-tested.
    /// </summary>
    public sealed class ProtectionSignals
    {
        public ProtectionSignals(string messageClass, string storedMessageClass = null,
                                 bool hasSmimePayload = false, bool hasUnsavedChanges = false)
        {
            MessageClass = messageClass;
            StoredMessageClass = storedMessageClass;
            HasSmimePayload = hasSmimePayload;
            HasUnsavedChanges = hasUnsavedChanges;
        }

        /// <summary>The class the opened item reports, or null when unreadable.</summary>
        public string MessageClass { get; }

        /// <summary>
        /// The class read from the folder's contents table (what is stored, not
        /// what an add-in shows in memory), or null when it could not be looked up.
        /// </summary>
        public string StoredMessageClass { get; }

        public bool HasSmimePayload { get; }

        public bool HasUnsavedChanges { get; }
    }

    /// <summary>
    /// The single definition of "protected" (encrypted or signed) mail: the
    /// add-in must never save or rewrite such an item (v2.11.0.0 A3, #62).
    /// <para>
    /// The live message class alone is not enough: with the smart card inserted,
    /// Stormshield decrypts an encrypted mail in memory when it is opened, and the
    /// opened item then reports a plain class. Saving it writes the decrypted view
    /// without its attachments (#62, confirmed on the target 2026-11-13). Any
    /// signal therefore makes the mail protected; a false positive only skips a
    /// convenience action, a false negative destroys the mail.
    /// </para>
    /// <c>PR_SECURITY_FLAGS</c> is recorded by the diagnostic trace but plays no
    /// part until it has been validated on the real target (decided 2026-11-05).
    /// </summary>
    public static class ProtectedMailPolicy
    {
        /// <summary>
        /// True for an S/MIME or secure message class. An unreadable class counts
        /// as protected: skipping a convenience action is recoverable, saving an
        /// encrypted item may not be.
        /// </summary>
        public static bool IsProtected(string messageClass)
        {
            if (string.IsNullOrWhiteSpace(messageClass)) return true;
            return IsProtectedClass(messageClass);
        }

        /// <summary>The first signal that makes the mail protected, or <see cref="ProtectionReason.None"/>.</summary>
        public static ProtectionReason Evaluate(ProtectionSignals signals)
        {
            if (signals == null) return ProtectionReason.UnreadableClass;
            if (string.IsNullOrWhiteSpace(signals.MessageClass)) return ProtectionReason.UnreadableClass;
            if (IsProtectedClass(signals.MessageClass)) return ProtectionReason.MessageClass;
            if (!string.IsNullOrWhiteSpace(signals.StoredMessageClass)
                && IsProtectedClass(signals.StoredMessageClass))
                return ProtectionReason.StoredMessageClass;
            if (signals.HasSmimePayload) return ProtectionReason.SmimePayload;
            if (signals.HasUnsavedChanges) return ProtectionReason.UnsavedChanges;
            return ProtectionReason.None;
        }

        /// <summary>
        /// True for the attachment names S/MIME uses for its payload: smime.p7m
        /// (enveloped/opaque-signed) and smime.p7s (detached signature). Other
        /// .p7m files (e.g. signed PDF documents) are ordinary attachments.
        /// </summary>
        public static bool IsSmimePayloadName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return false;
            string name = fileName.Trim();
            return string.Equals(name, "smime.p7m", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "smime.p7s", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsProtectedClass(string messageClass)
        {
            return messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase)
                || messageClass.StartsWith("IPM.Note.Secure", StringComparison.OrdinalIgnoreCase);
        }
    }
}
