using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class ExternalBannerStripperTests
    {
        // A realistic transport-rule banner: a single coloured table prepended
        // to the body, then the actual message.
        private const string Banner =
            "<table border=\"0\" cellpadding=\"0\" width=\"100%\" style=\"background:#FFFFCC\">" +
            "<tr><td><b>CAUTION:</b> This email originated from outside the organisation. " +
            "Do not click links or open attachments unless you recognise the sender.</td></tr>" +
            "</table>";

        private static string Body(string inner) =>
            "<html><head></head><body>" + inner + "</body></html>";

        [Fact]
        public void ExtractBannerBlock_returns_the_leading_table_verbatim()
        {
            var html = Body(Banner + "<p>Hello, here is the real message.</p>");

            ExternalBannerStripper.ExtractBannerBlock(html).Should().Be(Banner);
        }

        [Fact]
        public void ExtractBannerBlock_matches_the_correct_close_with_nested_tables()
        {
            // A banner table that itself contains a nested table must be captured
            // whole, not truncated at the inner </table>.
            string nested =
                "<table><tr><td>outer " +
                "<table><tr><td>inner</td></tr></table>" +
                " still outer</td></tr></table>";
            var html = Body(nested + "<div>body</div>");

            ExternalBannerStripper.ExtractBannerBlock(html).Should().Be(nested);
        }

        [Fact]
        public void ExtractBannerBlock_falls_back_to_a_div_when_there_is_no_table()
        {
            string div = "<div style=\"border:1px solid red\">External sender warning.</div>";
            var html = Body(div + "<p>message</p>");

            ExternalBannerStripper.ExtractBannerBlock(html).Should().Be(div);
        }

        [Fact]
        public void ExtractBannerBlock_skips_the_body_wrapper_and_takes_the_inner_table()
        {
            // Word/Outlook often wraps content in a div; the banner table inside
            // it is still what we want.
            var html =
                "<html><body><div class=\"WordSection1\">" + Banner +
                "<p>content</p></div></body></html>";

            ExternalBannerStripper.ExtractBannerBlock(html).Should().Be(Banner);
        }

        [Fact]
        public void ExtractBannerBlock_returns_null_when_there_is_no_block()
        {
            ExternalBannerStripper.ExtractBannerBlock("<html><body>just text</body></html>")
                                  .Should().BeNull();
            ExternalBannerStripper.ExtractBannerBlock(null).Should().BeNull();
            ExternalBannerStripper.ExtractBannerBlock("").Should().BeNull();
        }

        [Fact]
        public void Strip_removes_an_exact_banner_occurrence()
        {
            var html = Body(Banner + "<p>Real content stays.</p>");

            bool stripped;
            var result = ExternalBannerStripper.Strip(html, Banner, out stripped);

            stripped.Should().BeTrue();
            result.Should().NotContain("CAUTION");
            result.Should().Contain("Real content stays.");
        }

        [Fact]
        public void Strip_matches_despite_whitespace_reflow()
        {
            // The captured signature has different insignificant whitespace than
            // the target body (newlines/indentation), but the same structure.
            var signature =
                "<table border=\"0\">\r\n  <tr>\r\n    <td>External sender.</td>\r\n  </tr>\r\n</table>";
            var bodyVariant =
                Body("<table border=\"0\"><tr><td>External sender.</td></tr></table><p>hi</p>");

            bool stripped;
            var result = ExternalBannerStripper.Strip(bodyVariant, signature, out stripped);

            stripped.Should().BeTrue();
            result.Should().NotContain("External sender.");
            result.Should().Contain("hi");
        }

        [Fact]
        public void Strip_is_a_no_op_when_the_banner_is_absent()
        {
            var html = Body("<p>No banner here.</p>");

            bool stripped;
            var result = ExternalBannerStripper.Strip(html, Banner, out stripped);

            stripped.Should().BeFalse();
            result.Should().Be(html);
        }

        [Fact]
        public void Strip_tolerates_empty_inputs()
        {
            bool a, b, c;
            ExternalBannerStripper.Strip(null, Banner, out a).Should().BeNull();
            ExternalBannerStripper.Strip("<p>x</p>", null, out b).Should().Be("<p>x</p>");
            ExternalBannerStripper.Strip("<p>x</p>", "   ", out c).Should().Be("<p>x</p>");
            a.Should().BeFalse(); b.Should().BeFalse(); c.Should().BeFalse();
        }

        [Fact]
        public void Extract_then_strip_round_trips_on_a_reply_quoting_the_original()
        {
            // Learn from one mail, strip from a reply that quotes that mail
            // (banner embedded in the quoted history).
            var learnedFrom = Body(Banner + "<p>Original message.</p>");
            var signature = ExternalBannerStripper.ExtractBannerBlock(learnedFrom);

            var reply = Body(
                "<p>My reply.</p><hr>From: someone<br>" + Banner + "<p>Original message.</p>");

            bool stripped;
            var result = ExternalBannerStripper.Strip(reply, signature, out stripped);

            stripped.Should().BeTrue();
            result.Should().Contain("My reply.");
            result.Should().Contain("Original message.");
            result.Should().NotContain("CAUTION");
        }

        [Fact]
        public void ContainsBanner_reflects_presence()
        {
            ExternalBannerStripper.ContainsBanner(Body(Banner), Banner).Should().BeTrue();
            ExternalBannerStripper.ContainsBanner(Body("<p>nope</p>"), Banner).Should().BeFalse();
        }
    }
}
