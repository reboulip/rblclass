using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Whether a protected mail can take the re-encrypted filing path
    /// (v2.12.0.0, experimental). Only an encrypted mail readable decrypted in
    /// memory qualifies: signed-only mail has nothing to re-encrypt, and a mail
    /// still showing its S/MIME payload was not decrypted (smart card absent).
    /// Every other verdict sends the mail down the standard path, silently.
    /// </summary>
    public static class ReencryptEligibility
    {
        /// <summary>
        /// Null when the mail is eligible; otherwise why it is not
        /// (<see cref="ReencryptOutcome.Unresolved"/>,
        /// <see cref="ReencryptOutcome.SignedOnly"/>,
        /// <see cref="ReencryptOutcome.NotEncrypted"/> or
        /// <see cref="ReencryptOutcome.NotDecrypted"/>).
        /// </summary>
        /// <param name="state">The live and stored protection signals; null when the item did not resolve.</param>
        /// <param name="attachments">The attachments as the opened item shows them, or null when not read.</param>
        public static ReencryptOutcome? Evaluate(MailProtectionState state, IReadOnlyList<AttachmentInfo> attachments = null)
        {
            if (state == null || string.IsNullOrWhiteSpace(state.MessageClass))
                return ReencryptOutcome.Unresolved;

            if (IsSignedOnlyClass(state.MessageClass) || IsSignedOnlyClass(state.StoredMessageClass))
                return ReencryptOutcome.SignedOnly;

            if (!IsEncryptedClass(state.MessageClass) && !IsEncryptedClass(state.StoredMessageClass))
                return ReencryptOutcome.NotEncrypted;

            if (state.HasSmimePayload
                || (attachments != null && attachments.Any(a => a != null && ProtectedMailPolicy.IsSmimePayloadName(a.FileName))))
                return ReencryptOutcome.NotDecrypted;

            return null;
        }

        /// <summary>True when <see cref="Evaluate"/> finds the mail eligible.</summary>
        public static bool IsEligible(MailProtectionState state, IReadOnlyList<AttachmentInfo> attachments = null) =>
            Evaluate(state, attachments) == null;

        /// <summary>A multipart-signed or secure-sign class: signed, not encrypted.</summary>
        public static bool IsSignedOnlyClass(string messageClass) =>
            !string.IsNullOrWhiteSpace(messageClass)
            && (messageClass.StartsWith("IPM.Note.SMIME.MultipartSigned", StringComparison.OrdinalIgnoreCase)
                || messageClass.StartsWith("IPM.Note.Secure.Sign", StringComparison.OrdinalIgnoreCase));

        /// <summary>An S/MIME or secure class that is not signed-only.</summary>
        public static bool IsEncryptedClass(string messageClass) =>
            !string.IsNullOrWhiteSpace(messageClass)
            && !IsSignedOnlyClass(messageClass)
            && (messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase)
                || messageClass.StartsWith("IPM.Note.Secure", StringComparison.OrdinalIgnoreCase));
    }
}
