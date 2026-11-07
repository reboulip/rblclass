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
    }
}
