using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// The stored-state check of a re-encrypted copy (v2.14.0.1): read from the
    /// destination's contents table, never by opening the copy.
    /// </summary>
    public class EncryptedCopyCheckTests
    {
        private const long Payload = 40_000;

        [Fact]
        public void An_encrypted_row_with_an_attachment_and_room_for_the_payload_passes()
        {
            EncryptedCopyCheck.Evaluate(true, "IPM.Note.SMIME", true, 45_000, Payload)
                .Should().Be(EncryptedCopyProblem.None);
        }

        [Fact]
        public void The_class_is_compared_without_case_or_surrounding_spaces()
        {
            EncryptedCopyCheck.Evaluate(true, " ipm.note.smime ", true, 45_000, Payload)
                .Should().Be(EncryptedCopyProblem.None);
        }

        [Fact]
        public void A_missing_row_fails_first()
        {
            EncryptedCopyCheck.Evaluate(false, "IPM.Note.SMIME", true, 45_000, Payload)
                .Should().Be(EncryptedCopyProblem.RowNotFound);
        }

        [Theory]
        [InlineData("IPM.Note")]
        [InlineData("IPM.Post")]
        [InlineData("IPM.Note.SMIME.MultipartSigned")]
        [InlineData("")]
        [InlineData(null)]
        public void Any_other_stored_class_fails(string storedClass)
        {
            EncryptedCopyCheck.Evaluate(true, storedClass, true, 45_000, Payload)
                .Should().Be(EncryptedCopyProblem.NotEncryptedClass);
        }

        [Fact]
        public void A_row_without_an_attachment_fails()
        {
            EncryptedCopyCheck.Evaluate(true, "IPM.Note.SMIME", false, 45_000, Payload)
                .Should().Be(EncryptedCopyProblem.NoAttachment);
        }

        [Fact]
        public void A_row_smaller_than_the_payload_fails()
        {
            EncryptedCopyCheck.Evaluate(true, "IPM.Note.SMIME", true, Payload - 1, Payload)
                .Should().Be(EncryptedCopyProblem.TooSmall);
        }

        [Fact]
        public void A_row_exactly_the_payload_size_passes()
        {
            EncryptedCopyCheck.Evaluate(true, "IPM.Note.SMIME", true, Payload, Payload)
                .Should().Be(EncryptedCopyProblem.None);
        }

        [Fact]
        public void Unreadable_attachment_and_size_columns_are_not_failures()
        {
            EncryptedCopyCheck.Evaluate(true, "IPM.Note.SMIME", null, null, Payload)
                .Should().Be(EncryptedCopyProblem.None);
        }
    }
}
