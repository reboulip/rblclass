using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class ProtectedMailPolicyTests
    {
        [Theory]
        [InlineData("IPM.Note.SMIME")]
        [InlineData("IPM.Note.SMIME.MultipartSigned")]
        [InlineData("ipm.note.smime")]
        [InlineData("IPM.Note.Secure")]
        [InlineData("IPM.Note.Secure.Sign")]
        public void Secure_message_classes_are_protected(string messageClass)
        {
            ProtectedMailPolicy.IsProtected(messageClass).Should().BeTrue();
        }

        [Theory]
        [InlineData("IPM.Note")]
        [InlineData("IPM.NoteX")]
        [InlineData("IPM.Note.Rules.ReplyTemplate")]
        [InlineData("IPM.Schedule.Meeting.Request")]
        [InlineData("X.IPM.Note.SMIME")]
        public void Other_message_classes_are_not_protected(string messageClass)
        {
            ProtectedMailPolicy.IsProtected(messageClass).Should().BeFalse();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void An_unreadable_class_counts_as_protected(string messageClass)
        {
            ProtectedMailPolicy.IsProtected(messageClass).Should().BeTrue();
        }

        [Fact]
        public void Security_flags_do_not_decide_protection()
        {
            new MailProtectionState("IPM.Note", 1).IsProtected.Should().BeFalse();
            new MailProtectionState("IPM.Note.SMIME", null).IsProtected.Should().BeTrue();
        }

        [Fact]
        public void Plain_mail_with_no_other_signal_is_not_protected()
        {
            ProtectedMailPolicy.Evaluate(new ProtectionSignals("IPM.Note"))
                .Should().Be(ProtectionReason.None);
        }

        [Fact]
        public void The_live_class_wins_when_it_is_protected()
        {
            ProtectedMailPolicy.Evaluate(new ProtectionSignals("IPM.Note.SMIME", "IPM.Note.SMIME", true, true))
                .Should().Be(ProtectionReason.MessageClass);
        }

        [Theory]
        [InlineData(null)]
        [InlineData(" ")]
        public void An_unreadable_live_class_is_protected(string messageClass)
        {
            ProtectedMailPolicy.Evaluate(new ProtectionSignals(messageClass))
                .Should().Be(ProtectionReason.UnreadableClass);
        }

        [Fact]
        public void Missing_signals_count_as_unreadable()
        {
            ProtectedMailPolicy.Evaluate(null).Should().Be(ProtectionReason.UnreadableClass);
        }

        // #62: Stormshield, smart card inserted, decrypts the opened item in memory
        // and it reports IPM.Note; the folder still stores IPM.Note.SMIME.
        [Theory]
        [InlineData("IPM.Note.SMIME")]
        [InlineData("IPM.Note.SMIME.MultipartSigned")]
        [InlineData("IPM.Note.Secure")]
        public void A_protected_stored_class_protects_a_mail_decrypted_in_memory(string stored)
        {
            ProtectedMailPolicy.Evaluate(new ProtectionSignals("IPM.Note", stored))
                .Should().Be(ProtectionReason.StoredMessageClass);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("IPM.Note")]
        public void An_unknown_or_plain_stored_class_adds_no_protection(string stored)
        {
            ProtectedMailPolicy.Evaluate(new ProtectionSignals("IPM.Note", stored))
                .Should().Be(ProtectionReason.None);
        }

        [Fact]
        public void An_smime_payload_attachment_protects_the_mail()
        {
            ProtectedMailPolicy.Evaluate(new ProtectionSignals("IPM.Note", null, hasSmimePayload: true))
                .Should().Be(ProtectionReason.SmimePayload);
        }

        [Fact]
        public void Unsaved_changes_on_open_protect_the_mail()
        {
            ProtectedMailPolicy.Evaluate(new ProtectionSignals("IPM.Note", null, hasUnsavedChanges: true))
                .Should().Be(ProtectionReason.UnsavedChanges);
        }

        [Fact]
        public void The_stored_class_is_reported_before_weaker_signals()
        {
            ProtectedMailPolicy.Evaluate(new ProtectionSignals("IPM.Note", "IPM.Note.SMIME", true, true))
                .Should().Be(ProtectionReason.StoredMessageClass);
        }

        [Theory]
        [InlineData("smime.p7m", true)]
        [InlineData("SMIME.P7M", true)]
        [InlineData("smime.p7s", true)]
        [InlineData(" smime.p7m ", true)]
        [InlineData("contrat.pdf.p7m", false)]
        [InlineData("smime.p7m.txt", false)]
        [InlineData("report.pdf", false)]
        [InlineData(null, false)]
        [InlineData("", false)]
        public void Only_the_smime_part_names_are_payload(string fileName, bool expected)
        {
            ProtectedMailPolicy.IsSmimePayloadName(fileName).Should().Be(expected);
        }

        [Fact]
        public void The_state_reports_protection_from_the_stored_class()
        {
            var state = new MailProtectionState("IPM.Note", 0, storedMessageClass: "IPM.Note.SMIME");
            state.IsProtected.Should().BeTrue();
            state.Reason.Should().Be(ProtectionReason.StoredMessageClass);
        }
    }
}
