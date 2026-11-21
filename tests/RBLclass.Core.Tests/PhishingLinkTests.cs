using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class PhishingLinkTests
    {
        private const string Signature =
            "<table><tr><td><p>This message is from an <b>EXTERNAL SENDER</b> and could be malicious</p>" +
            "<p>Ce message provient d&#8217;un <b>EXPEDITEUR EXTERNE</b> et pourrait &ecirc;tre malveillant</p></td></tr></table>";

        // ---- ExternalBannerStripper.ContainsBannerText -------------------

        [Theory]
        [InlineData("Ce message provient d’un EXPEDITEUR EXTERNE et pourrait être malveillant")]
        [InlineData("Ce message provient d'un EXPEDITEUR EXTERNE et pourrait être malveillant")]
        [InlineData("Ce message provient d&#8217;un EXPEDITEUR EXTERNE et pourrait être malveillant")]
        public void ContainsBannerText_ignores_quote_style_and_entities(string rendered)
        {
            var text = "This message is from an EXTERNAL SENDER and could be malicious\v" + rendered + "\r";

            ExternalBannerStripper.ContainsBannerText(text, Signature).Should().BeTrue();
        }

        [Fact]
        public void ContainsBannerText_matches_word_text_with_vertical_tab_and_carriage_return_breaks()
        {
            var text = "Hello\r\rThis message is from an\vEXTERNAL SENDER\r and could be malicious\vCe message provient d’un EXPEDITEUR EXTERNE\r et pourrait être malveillant\r\rBody";

            ExternalBannerStripper.ContainsBannerText(text, Signature).Should().BeTrue();
        }

        [Fact]
        public void ContainsBannerText_does_not_match_when_the_last_word_is_missing()
        {
            var text = "This message is from an EXTERNAL SENDER and could be malicious Ce message provient d’un EXPEDITEUR EXTERNE et pourrait être";

            ExternalBannerStripper.ContainsBannerText(text, Signature).Should().BeFalse();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void ContainsBannerText_is_false_for_a_blank_signature(string signature)
        {
            ExternalBannerStripper.ContainsBannerText("anything at all here", signature).Should().BeFalse();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ContainsBannerText_is_false_for_a_blank_text(string text)
        {
            ExternalBannerStripper.ContainsBannerText(text, Signature).Should().BeFalse();
        }

        [Fact]
        public void ContainsBannerText_is_false_for_a_trivial_signature()
        {
            ExternalBannerStripper.ContainsBannerText("External sender", "<b>External</b>").Should().BeFalse();
        }

        // ---- PhishingLinkFilter ------------------------------------------

        [Theory]
        [InlineData(null, false)]
        [InlineData("", false)]
        [InlineData("   ", false)]
        [InlineData("#section", false)]
        [InlineData("  #top", false)]
        [InlineData("mailto:someone@example.com", false)]
        [InlineData("MAILTO:someone@example.com", false)]
        [InlineData("http://example.com/a", true)]
        [InlineData("https://example.com/a?b=c", true)]
        [InlineData("ftp://example.com/file", true)]
        [InlineData("file:///C:/temp/x.exe", true)]
        [InlineData("custom-scheme://thing", true)]
        public void IsWarnable_covers_each_target_kind(string url, bool expected)
        {
            PhishingLinkFilter.IsWarnable(url).Should().Be(expected);
        }

        [Theory]
        [InlineData("https://example.com/secret/path?token=abc", "https://example.com")]
        [InlineData("http://Sub.Example.org:8080/x", "http://sub.example.org")]
        [InlineData("not a url", "(unparsed)")]
        [InlineData("", "(unparsed)")]
        [InlineData(null, "(unparsed)")]
        public void HostForLog_never_includes_the_path_or_query(string url, string expected)
        {
            PhishingLinkFilter.HostForLog(url).Should().Be(expected);
        }

        // ---- PhishingWarningPolicy.ShouldWarnLink ------------------------

        private const string Rendered =
            "Hello This message is from an EXTERNAL SENDER and could be malicious Ce message provient d’un EXPEDITEUR EXTERNE et pourrait être malveillant";

        [Fact]
        public void ShouldWarnLink_is_true_when_everything_lines_up()
        {
            PhishingWarningPolicy.ShouldWarnLink(true, Signature, Rendered, "https://example.com").Should().BeTrue();
        }

        [Fact]
        public void ShouldWarnLink_is_false_when_the_setting_is_off()
        {
            PhishingWarningPolicy.ShouldWarnLink(false, Signature, Rendered, "https://example.com").Should().BeFalse();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ShouldWarnLink_is_false_without_a_learned_signature(string signature)
        {
            PhishingWarningPolicy.ShouldWarnLink(true, signature, Rendered, "https://example.com").Should().BeFalse();
        }

        [Theory]
        [InlineData("#anchor")]
        [InlineData("mailto:a@b.c")]
        [InlineData("")]
        public void ShouldWarnLink_is_false_for_targets_that_do_not_warrant_a_warning(string url)
        {
            PhishingWarningPolicy.ShouldWarnLink(true, Signature, Rendered, url).Should().BeFalse();
        }

        [Fact]
        public void ShouldWarnLink_is_false_when_the_rendered_text_has_no_banner()
        {
            PhishingWarningPolicy.ShouldWarnLink(true, Signature, "Just a normal mail body", "https://example.com").Should().BeFalse();
        }

        // ---- LinkWarningMessage ------------------------------------------

        [Fact]
        public void DisplayedTextDiffers_is_true_when_the_mail_shows_something_else()
        {
            var message = new LinkWarningMessage(null, "https://evil.example.net/login", "https://bank.example.com");

            message.DisplayedTextDiffers.Should().BeTrue();
        }

        [Theory]
        [InlineData("https://example.com", "https://example.com")]
        [InlineData("https://example.com/", "https://example.com")]
        [InlineData("https://EXAMPLE.com", " https://example.com/ ")]
        [InlineData("https://example.com", "")]
        [InlineData("https://example.com", null)]
        public void DisplayedTextDiffers_is_false_when_the_text_is_the_destination_or_empty(string url, string shown)
        {
            new LinkWarningMessage(null, url, shown).DisplayedTextDiffers.Should().BeFalse();
        }

        [Fact]
        public void A_missing_sender_becomes_unknown()
        {
            new LinkWarningMessage(null, "https://x.example", null).Sender.IsKnown.Should().BeFalse();
        }
    }
}
