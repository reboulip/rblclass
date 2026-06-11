using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Result of widening a mail to its conversation: the siblings that can be
    /// filed/moved, plus the subjects of any that were deliberately left in
    /// place because they are encrypted/signed (S/MIME, e.g. Stormshield).
    /// Those can't be safely processed when the encryption provider is inactive,
    /// so they are skipped and reported rather than silently dropped.
    /// </summary>
    public sealed class ConversationSiblings
    {
        public ConversationSiblings(IReadOnlyList<MailItemRef> processable,
                                    IReadOnlyList<string> skippedEncryptedSubjects)
        {
            Processable = processable ?? new MailItemRef[0];
            SkippedEncryptedSubjects = skippedEncryptedSubjects ?? new string[0];
        }

        /// <summary>Conversation siblings (in scope) that can be filed/moved.</summary>
        public IReadOnlyList<MailItemRef> Processable { get; }

        /// <summary>Subjects of in-scope siblings skipped because they are encrypted/signed.</summary>
        public IReadOnlyList<string> SkippedEncryptedSubjects { get; }

        public static ConversationSiblings Empty { get; } =
            new ConversationSiblings(new MailItemRef[0], new string[0]);
    }
}
