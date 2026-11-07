using System;

namespace RBLclass.Core
{
    /// <summary>
    /// The single definition of "protected" (encrypted or signed) mail: the
    /// add-in must never save or rewrite such an item (v2.11.0.0 A3, #62).
    /// Decided on the message class only. <c>PR_SECURITY_FLAGS</c> is recorded by
    /// the diagnostic trace but deliberately plays no part until it has been
    /// validated on the real target.
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

            return messageClass.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase)
                || messageClass.StartsWith("IPM.Note.Secure", StringComparison.OrdinalIgnoreCase);
        }
    }
}
