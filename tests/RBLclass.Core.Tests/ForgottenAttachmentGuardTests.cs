using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class ForgottenAttachmentGuardTests
    {
        private static readonly string[] Keywords = { "attach", "enclos", "joint", "PJ" };

        [Fact]
        public void Warns_when_no_attachment_and_body_mentions_a_keyword()
        {
            var sut = new ForgottenAttachmentGuard();

            sut.ShouldWarn("Please find the document attached.", attachmentCount: 0, Keywords).Should().BeTrue();
        }

        [Fact]
        public void Matches_keywords_case_insensitively()
        {
            var sut = new ForgottenAttachmentGuard();

            sut.ShouldWarn("Voir le fichier ci-JOINT.", attachmentCount: 0, Keywords).Should().BeTrue();
        }

        [Fact]
        public void Does_not_warn_when_an_attachment_is_present()
        {
            var sut = new ForgottenAttachmentGuard();

            sut.ShouldWarn("Please find the document attached.", attachmentCount: 1, Keywords).Should().BeFalse();
        }

        [Fact]
        public void Does_not_warn_when_the_body_mentions_no_keyword()
        {
            var sut = new ForgottenAttachmentGuard();

            sut.ShouldWarn("Let's catch up tomorrow.", attachmentCount: 0, Keywords).Should().BeFalse();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void Does_not_warn_on_an_empty_body(string body)
        {
            var sut = new ForgottenAttachmentGuard();

            sut.ShouldWarn(body, attachmentCount: 0, Keywords).Should().BeFalse();
        }

        [Fact]
        public void Does_not_warn_when_no_keywords_are_configured()
        {
            var sut = new ForgottenAttachmentGuard();

            sut.ShouldWarn("Please find the document attached.", attachmentCount: 0, new string[0]).Should().BeFalse();
        }
    }
}
