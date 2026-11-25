using System;

namespace RBLclass.Core
{
    /// <summary>Why a re-encrypted copy failed its stored-state check (v2.14.0.1).</summary>
    public enum EncryptedCopyProblem
    {
        /// <summary>The copy is stored as an encrypted item carrying its payload.</summary>
        None,

        /// <summary>No row of the destination's contents table is the copy.</summary>
        RowNotFound,

        /// <summary>The copy is not stored with the <c>IPM.Note.SMIME</c> class.</summary>
        NotEncryptedClass,

        /// <summary>The copy is stored without any attachment: the payload is missing.</summary>
        NoAttachment,

        /// <summary>The stored copy is smaller than the payload it should carry.</summary>
        TooSmall,
    }

    /// <summary>
    /// Checks a re-encrypted copy from what its folder's contents table shows: the
    /// stored class, whether it has an attachment and its size (v2.14.0.1). The
    /// copy is never opened: on the target, Stormshield decrypts in memory an item
    /// whose <c>smime.p7m</c> part is read through the object model, so reading the
    /// payload back to compare it can never succeed there.
    /// </summary>
    public static class EncryptedCopyCheck
    {
        /// <summary>The class a re-encrypted copy is stored with.</summary>
        public const string EncryptedClass = "IPM.Note.SMIME";

        /// <param name="rowFound">A row of the destination's contents table is the copy.</param>
        /// <param name="storedClass">The row's message class.</param>
        /// <param name="hasAttachment">The row's <c>PR_HASATTACH</c>, or null when the column could not be read.</param>
        /// <param name="storedSize">The row's <c>PR_MESSAGE_SIZE</c>, or null when the column could not be read.</param>
        /// <param name="payloadLength">The byte count of the <c>smime.p7m</c> payload attached to the copy.</param>
        public static EncryptedCopyProblem Evaluate(bool rowFound, string storedClass, bool? hasAttachment,
                                                    long? storedSize, long payloadLength)
        {
            if (!rowFound) return EncryptedCopyProblem.RowNotFound;

            // Exactly the class the copy is given: a signed-only class
            // (IPM.Note.SMIME.MultipartSigned) or a plain IPM.Note is a failure.
            if (!string.Equals(storedClass?.Trim(), EncryptedClass, StringComparison.OrdinalIgnoreCase))
                return EncryptedCopyProblem.NotEncryptedClass;

            // An unreadable column is not evidence of a problem; a store saying
            // "no attachment" or "smaller than the payload" is.
            if (hasAttachment == false) return EncryptedCopyProblem.NoAttachment;
            if (storedSize.HasValue && storedSize.Value < payloadLength) return EncryptedCopyProblem.TooSmall;

            return EncryptedCopyProblem.None;
        }
    }
}
