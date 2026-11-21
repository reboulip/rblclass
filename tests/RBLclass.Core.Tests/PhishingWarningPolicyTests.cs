using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class PhishingWarningPolicyTests
    {
        private const string Banner = "<table><tr><td>Ce message provient d'un expediteur externe</td></tr></table>";
        private static readonly string BodyWithBanner = "<html><body>" + Banner + "<p>Hello</p></body></html>";
        private const string BodyWithout = "<html><body><p>Hello</p></body></html>";

        [Fact]
        public void Warns_when_the_setting_is_on_and_the_banner_is_present()
        {
            PhishingWarningPolicy.ShouldWarnAttachment(true, Banner, BodyWithBanner).Should().BeTrue();
        }

        [Fact]
        public void Does_not_warn_when_the_setting_is_off()
        {
            PhishingWarningPolicy.ShouldWarnAttachment(false, Banner, BodyWithBanner).Should().BeFalse();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void Does_not_warn_when_no_signature_was_learned(string signature)
        {
            PhishingWarningPolicy.ShouldWarnAttachment(true, signature, BodyWithBanner).Should().BeFalse();
        }

        [Fact]
        public void Does_not_warn_when_the_body_has_no_banner()
        {
            PhishingWarningPolicy.ShouldWarnAttachment(true, Banner, BodyWithout).Should().BeFalse();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void Does_not_warn_for_an_empty_body(string body)
        {
            PhishingWarningPolicy.ShouldWarnAttachment(true, Banner, body).Should().BeFalse();
        }

        [Fact]
        public void Banner_match_tolerates_whitespace_run_differences()
        {
            var signature = "<table>\n<tr>\n<td>Ce message provient d'un expediteur externe</td>\n</tr>\n</table>";
            var body = "<html><body><table>\r\n\r\n<tr>\r\n\r\n<td>Ce message provient d'un expediteur externe</td>\r\n\r\n</tr>\r\n</table></body></html>";

            PhishingWarningPolicy.ShouldWarnAttachment(true, signature, body).Should().BeTrue();
        }

        [Fact]
        public void Message_defaults_to_an_unknown_sender_and_no_file_name()
        {
            var message = new PhishingWarningMessage(null, null);

            message.Sender.IsKnown.Should().BeFalse();
            message.HasAttachmentFileName.Should().BeFalse();
        }

        [Fact]
        public void Message_carries_the_sender_and_file_name()
        {
            var message = new PhishingWarningMessage(new SenderInfo("Alice", "alice@example.com"), "invoice.pdf");

            message.Sender.SmtpAddress.Should().Be("alice@example.com");
            message.AttachmentFileName.Should().Be("invoice.pdf");
            message.HasAttachmentFileName.Should().BeTrue();
        }
    }
}
