using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class ReencryptEligibilityTests
    {
        private static MailProtectionState State(string live, string stored = null, bool payload = false, bool unsaved = false) =>
            new MailProtectionState(live, null, stored, payload, unsaved);

        [Theory]
        [InlineData("IPM.Note", "IPM.Note.SMIME")]          // Stormshield: decrypted in memory
        [InlineData("IPM.Note.SMIME", "IPM.Note.SMIME")]    // Outlook's own S/MIME, payload hidden
        [InlineData("IPM.Note.SMIME", null)]                // stored class not found
        [InlineData("IPM.Note.Secure", null)]
        public void Encrypted_mail_readable_decrypted_is_eligible(string live, string stored)
        {
            ReencryptEligibility.Evaluate(State(live, stored)).Should().BeNull();
            ReencryptEligibility.IsEligible(State(live, stored)).Should().BeTrue();
        }

        [Theory]
        [InlineData("IPM.Note.SMIME.MultipartSigned", null)]
        [InlineData("IPM.Note", "IPM.Note.SMIME.MultipartSigned")]
        [InlineData("IPM.Note.Secure.Sign", null)]
        [InlineData("IPM.Note.SMIME", "ipm.note.secure.sign")]
        public void Signed_only_mail_is_not_eligible(string live, string stored)
        {
            ReencryptEligibility.Evaluate(State(live, stored)).Should().Be(ReencryptOutcome.SignedOnly);
        }

        [Fact]
        public void A_visible_payload_flag_means_not_decrypted()
        {
            ReencryptEligibility.Evaluate(State("IPM.Note.SMIME", payload: true)).Should().Be(ReencryptOutcome.NotDecrypted);
        }

        [Fact]
        public void A_visible_payload_attachment_means_not_decrypted()
        {
            var attachments = new[] { new AttachmentInfo(1, " SMIME.P7M ", 10) };

            ReencryptEligibility.Evaluate(State("IPM.Note.SMIME"), attachments).Should().Be(ReencryptOutcome.NotDecrypted);
        }

        [Fact]
        public void Another_p7m_attachment_does_not_block()
        {
            var attachments = new[] { new AttachmentInfo(1, "contrat.pdf.p7m", 10) };

            ReencryptEligibility.Evaluate(State("IPM.Note", "IPM.Note.SMIME"), attachments).Should().BeNull();
        }

        [Fact]
        public void Mail_protected_only_by_unsaved_changes_is_not_encrypted()
        {
            ReencryptEligibility.Evaluate(State("IPM.Note", "IPM.Note", unsaved: true)).Should().Be(ReencryptOutcome.NotEncrypted);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void An_unreadable_class_is_unresolved(string live)
        {
            ReencryptEligibility.Evaluate(State(live, "IPM.Note.SMIME")).Should().Be(ReencryptOutcome.Unresolved);
        }

        [Fact]
        public void A_missing_state_is_unresolved()
        {
            ReencryptEligibility.Evaluate(null).Should().Be(ReencryptOutcome.Unresolved);
        }
    }
}
