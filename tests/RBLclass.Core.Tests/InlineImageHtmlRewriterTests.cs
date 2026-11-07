using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class InlineImageHtmlRewriterTests
    {
        private const string Placeholder = "[Image removed]";

        private static InlineImageRewriteResult Rewrite(string html, params string[] ids) =>
            InlineImageHtmlRewriter.Rewrite(html, ids, Placeholder);

        [Fact]
        public void A_double_quoted_cid_image_becomes_the_placeholder()
        {
            var r = Rewrite("<p>a<img src=\"cid:logo@x\" width=\"10\">b</p>", "logo@x");

            r.Html.Should().Be("<p>a[Image removed]b</p>");
            r.ReplacedContentIds.Should().Equal("logo@x");
        }

        [Theory]
        [InlineData("<img src='cid:logo@x'>")]
        [InlineData("<img src=cid:logo@x>")]
        [InlineData("<IMG SRC=\"CID:logo@x\">")]
        [InlineData("<img alt=\"a>b\" src=\"cid:logo@x\">")]
        [InlineData("<img src=\"cid:logo%40x\">")]
        [InlineData("<img\nsrc=\"cid:logo@x\"/>")]
        public void Every_spelling_of_the_reference_is_replaced(string tag)
        {
            Rewrite("x" + tag + "y", "logo@x").Html.Should().Be("x[Image removed]y");
        }

        [Fact]
        public void An_html_encoded_content_id_matches()
        {
            Rewrite("<img src=\"cid:a&amp;b\">", "a&b").Html.Should().Be("[Image removed]");
        }

        [Fact]
        public void A_content_id_given_in_angle_brackets_matches()
        {
            Rewrite("<img src=\"cid:logo@x\">", "<logo@x>").Html.Should().Be("[Image removed]");
        }

        [Fact]
        public void A_repeated_image_is_replaced_everywhere_and_reported_once()
        {
            var r = Rewrite("<img src=\"cid:a\"><img src=\"cid:a\">", "a");

            r.Html.Should().Be("[Image removed][Image removed]");
            r.ReplacedContentIds.Should().Equal("a");
        }

        [Fact]
        public void Other_images_are_left_alone()
        {
            const string html = "<img src=\"cid:other\"><img src=\"http://x/a.png\"><img src=\"data:image/png;base64,AA\"><image src=\"cid:a\">";

            var r = Rewrite(html, "a");

            r.Html.Should().Be(html);
            r.ReplacedContentIds.Should().BeEmpty();
        }

        [Fact]
        public void The_word_vml_form_is_removed_and_its_fallback_unwrapped()
        {
            const string html =
                "<p><!--[if gte vml 1]><v:shape><v:imagedata src=\"cid:img1@x\" o:title=\"\"/></v:shape><![endif]-->" +
                "<![if !vml]><img width=\"20\" src=\"cid:img1@x\" v:shapes=\"Image_1\"><![endif]></p>";

            var r = Rewrite(html, "img1@x");

            r.Html.Should().Be("<p>[Image removed]</p>");
            r.ReplacedContentIds.Should().Equal("img1@x");
        }

        [Fact]
        public void A_vml_block_for_another_image_is_kept()
        {
            const string html = "<!--[if gte vml 1]><v:imagedata src=\"cid:keep\"/><![endif]--><![if !vml]><img src=\"cid:keep\"><![endif]>";

            Rewrite(html, "gone").Html.Should().Be(html);
        }

        [Fact]
        public void The_placeholder_is_html_encoded()
        {
            InlineImageHtmlRewriter.Rewrite("<img src=\"cid:a\">", new[] { "a" }, "<b>&")
                .Html.Should().Be("&lt;b&gt;&amp;");
        }

        [Fact]
        public void No_match_leaves_the_body_unchanged()
        {
            var r = Rewrite("<p>hello</p>", "a");

            r.Html.Should().Be("<p>hello</p>");
            r.ReplacedContentIds.Should().BeEmpty();
        }

        [Fact]
        public void A_missing_body_or_no_ids_is_handled()
        {
            Rewrite(null, "a").Html.Should().BeEmpty();
            InlineImageHtmlRewriter.Rewrite("<img src=\"cid:a\">", null, Placeholder).Html.Should().Be("<img src=\"cid:a\">");
        }

        [Fact]
        public void An_unterminated_tag_is_left_alone()
        {
            Rewrite("<p><img src=\"cid:a\"", "a").Html.Should().Be("<p><img src=\"cid:a\"");
        }
    }
}
