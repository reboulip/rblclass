using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class MimeEntityBuilderTests
    {
        private static readonly byte[] Pdf = Encoding.ASCII.GetBytes("%PDF-1.4 fake");
        private static readonly byte[] Png = { 0x89, 0x50, 0x4E, 0x47, 1, 2, 3 };

        private static MailPart Attachment(string name = "report.pdf", string type = "application/pdf", int id = 1) =>
            new MailPart(id, name, type, Pdf);

        private static MailPart InlineImage(string cid = "logo@x", int id = 2) =>
            new MailPart(id, "logo.png", "image/png", Png, isInline: true, contentId: cid);

        private static MimeNode Parse(MailContent content, MimeEntityBuilder builder = null) =>
            MimeTestReader.Parse((builder ?? new MimeEntityBuilder()).Build(content));

        // ---- structure ----------------------------------------------------------

        [Fact]
        public void Html_only_is_a_single_html_part_with_crlf_lines()
        {
            var root = Parse(new MailContent("<p>Bonjour\nété</p>", null, null));

            root.MediaType.Should().Be("text/html");
            root.TypeParams["charset"].Should().Be("utf-8");
            root.Headers["MIME-Version"].Should().Be("1.0");
            root.Text.Should().Be("<p>Bonjour\r\nété</p>");
        }

        [Fact]
        public void Plain_only_is_a_single_plain_part()
        {
            var root = Parse(new MailContent(null, "Hello", null));

            root.MediaType.Should().Be("text/plain");
            root.Text.Should().Be("Hello");
        }

        [Fact]
        public void Both_bodies_make_an_alternative_with_plain_first()
        {
            var root = Parse(new MailContent("<p>x</p>", "x", null));

            root.MediaType.Should().Be("multipart/alternative");
            root.Children.Select(c => c.MediaType).Should().Equal("text/plain", "text/html");
        }

        [Fact]
        public void Html_with_an_inline_image_makes_a_related_part()
        {
            var root = Parse(new MailContent("<img src=\"cid:logo@x\">", null, new[] { InlineImage() }));

            root.MediaType.Should().Be("multipart/related");
            root.TypeParams["type"].Should().Be("text/html");
            root.Children[0].MediaType.Should().Be("text/html");
            var image = root.Children[1];
            image.MediaType.Should().Be("image/png");
            image.ContentId.Should().Be("<logo@x>");
            image.Disposition.Should().Be("inline");
            image.Body.Should().Equal(Png);
        }

        [Fact]
        public void Html_with_an_attachment_makes_a_mixed_part()
        {
            var root = Parse(new MailContent("<p>x</p>", null, new[] { Attachment() }));

            root.MediaType.Should().Be("multipart/mixed");
            root.Children[0].MediaType.Should().Be("text/html");
            var file = root.Children[1];
            file.MediaType.Should().Be("application/pdf");
            file.Disposition.Should().Be("attachment");
            file.FileName.Should().Be("report.pdf");
            file.Name.Should().Be("report.pdf");
            file.Body.Should().Equal(Pdf);
        }

        [Fact]
        public void Everything_nests_mixed_related_alternative()
        {
            var root = Parse(new MailContent("<img src=\"cid:logo@x\">", "plain", new[] { InlineImage(), Attachment() }));

            root.MediaType.Should().Be("multipart/mixed");
            var related = root.Children[0];
            related.MediaType.Should().Be("multipart/related");
            related.TypeParams["type"].Should().Be("multipart/alternative");
            related.Children[0].MediaType.Should().Be("multipart/alternative");
            related.Children[0].Children.Select(c => c.MediaType).Should().Equal("text/plain", "text/html");
            related.Children[1].ContentId.Should().Be("<logo@x>");
            root.Children[1].FileName.Should().Be("report.pdf");
        }

        [Fact]
        public void An_inline_image_without_an_html_body_goes_to_mixed()
        {
            var root = Parse(new MailContent(null, "plain", new[] { InlineImage() }));

            root.MediaType.Should().Be("multipart/mixed");
            root.Children[1].Disposition.Should().Be("inline");
            root.Children[1].ContentId.Should().Be("<logo@x>");
        }

        [Fact]
        public void An_inline_part_without_a_content_id_goes_to_mixed_as_inline()
        {
            var hidden = new MailPart(3, "hidden.bin", null, Pdf, isInline: true, contentId: null);

            var root = Parse(new MailContent("<p>x</p>", null, new[] { hidden }));

            root.MediaType.Should().Be("multipart/mixed");
            root.Children[1].Disposition.Should().Be("inline");
            root.Children[1].ContentId.Should().BeNull();
            root.Children[1].MediaType.Should().Be("application/octet-stream");
        }

        [Fact]
        public void No_body_and_no_parts_is_an_empty_plain_part()
        {
            var root = Parse(new MailContent(null, null, null));

            root.MediaType.Should().Be("text/plain");
            root.Body.Should().BeEmpty();
        }

        [Fact]
        public void An_empty_body_inside_a_multipart_is_still_well_formed()
        {
            var root = Parse(new MailContent(string.Empty, null, new[] { Attachment() }));

            root.Children[0].MediaType.Should().Be("text/plain");
            root.Children[0].Body.Should().BeEmpty();
            root.Children[1].Body.Should().Equal(Pdf);
        }

        // ---- boundaries ---------------------------------------------------------

        [Fact]
        public void A_boundary_found_in_header_text_is_rejected()
        {
            var queue = new Queue<string>(new[] { "=_RBL_dup", "=_RBL_ok" });
            var builder = new MimeEntityBuilder(() => queue.Dequeue());
            var image = InlineImage(cid: "x=_RBL_dup@y");

            var root = Parse(new MailContent("<img src=\"cid:x=_RBL_dup@y\">", null, new[] { image }), builder);

            root.TypeParams["boundary"].Should().Be("=_RBL_ok");
            root.Children[1].Body.Should().Equal(Png);
        }

        [Fact]
        public void A_boundary_that_contains_another_is_rejected()
        {
            var queue = new Queue<string>(new[] { "=_A", "=_A1", "=_B" });
            var builder = new MimeEntityBuilder(() => queue.Dequeue());

            var root = Parse(new MailContent("<img src=\"cid:logo@x\">", null, new[] { InlineImage(), Attachment() }), builder);

            root.TypeParams["boundary"].Should().Be("=_A");
            root.Children[0].TypeParams["boundary"].Should().Be("=_B");
        }

        [Fact]
        public void Running_out_of_boundaries_throws()
        {
            var builder = new MimeEntityBuilder(() => "logo");

            Action build = () => builder.Build(new MailContent("<img>", null, new[] { InlineImage(cid: "logo") }));

            build.Should().Throw<InvalidOperationException>();
        }

        [Fact]
        public void Two_builds_use_different_boundaries()
        {
            var content = new MailContent("<p>x</p>", null, new[] { Attachment() });

            Parse(content).TypeParams["boundary"].Should().NotBe(Parse(content).TypeParams["boundary"]);
        }

        // ---- line rules ---------------------------------------------------------

        [Fact]
        public void Every_line_ends_with_crlf_and_respects_the_length_limits()
        {
            var big = new MailPart(4, new string('é', 150) + ".txt", "text/plain", new byte[10000]);
            var bytes = new MimeEntityBuilder().Build(
                new MailContent("<p>" + new string('x', 5000) + "</p>", "plain", new[] { big, InlineImage(), Attachment() }));
            string text = Encoding.ASCII.GetString(bytes);

            text.Should().EndWith("\r\n");
            var lines = text.Split(new[] { "\r\n" }, StringSplitOptions.None);
            lines.Should().OnlyContain(l => !l.Contains('\r') && !l.Contains('\n'));
            lines.Should().OnlyContain(l => l.Length <= 998);
            lines.Where(l => l.Length > 0 && l.All(c => char.IsLetterOrDigit(c) || c == '+' || c == '/' || c == '='))
                 .Should().OnlyContain(l => l.Length <= 76);
            bytes.Should().OnlyContain(b => b < 0x80);
        }

        // ---- names & sanitising -------------------------------------------------

        [Fact]
        public void A_long_non_ascii_file_name_round_trips()
        {
            string name = string.Concat(Enumerable.Repeat("Données été Prüfbericht 😀 ", 9)) + ".pdf";
            name.Length.Should().BeGreaterThan(200);

            var root = Parse(new MailContent("<p>x</p>", null, new[] { Attachment(name) }));

            root.Children[1].FileName.Should().Be(name);
            root.Children[1].Name.Should().Be(name);
        }

        [Fact]
        public void A_short_accented_name_uses_the_single_extended_form()
        {
            var bytes = new MimeEntityBuilder().Build(new MailContent("<p>x</p>", null, new[] { Attachment("été.pdf") }));
            string text = Encoding.ASCII.GetString(bytes);

            text.Should().Contain("filename*=utf-8''%C3%A9t%C3%A9.pdf");
            MimeTestReader.Parse(bytes).Children[1].FileName.Should().Be("été.pdf");
        }

        [Fact]
        public void A_simple_ascii_name_is_written_plainly()
        {
            string text = Encoding.ASCII.GetString(new MimeEntityBuilder().Build(
                new MailContent("<p>x</p>", null, new[] { Attachment("report.pdf") })));

            text.Should().Contain("filename=\"report.pdf\"").And.Contain("name=\"report.pdf\"");
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not a type")]
        [InlineData("text/plain\r\nX-Injected: yes")]
        public void An_invalid_content_type_becomes_octet_stream(string type)
        {
            var bytes = new MimeEntityBuilder().Build(new MailContent("<p>x</p>", null, new[] { Attachment(type: type) }));
            string text = Encoding.ASCII.GetString(bytes);

            MimeTestReader.Parse(bytes).Children[1].MediaType.Should().Be("application/octet-stream");
            text.Should().NotContain("\r\nX-Injected");
        }

        [Fact]
        public void Content_type_parameters_and_case_are_dropped()
        {
            var root = Parse(new MailContent("<p>x</p>", null, new[] { Attachment(type: "Application/PDF; name=x") }));

            root.Children[1].MediaType.Should().Be("application/pdf");
        }

        [Fact]
        public void Control_characters_are_stripped_from_names_and_content_ids()
        {
            var image = new MailPart(2, "lo\r\ngo.png", "image/png", Png, isInline: true, contentId: "lo\r\nX-Evil: 1go@x");
            var bytes = new MimeEntityBuilder().Build(new MailContent("<img>", null, new[] { image }));
            string text = Encoding.ASCII.GetString(bytes);

            text.Should().NotContain("\r\nX-Evil");
            var root = MimeTestReader.Parse(bytes);
            root.Children[1].ContentId.Should().Be("<loX-Evil: 1go@x>");
            root.Children[1].FileName.Should().Be("logo.png");
        }

        // ---- meta charset -------------------------------------------------------

        [Theory]
        [InlineData("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=windows-1252\">",
                    "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">")]
        [InlineData("<META charset=\"iso-8859-1\">", "<META charset=\"utf-8\">")]
        [InlineData("<meta charset=windows-1252>", "<meta charset=utf-8>")]
        [InlineData("<p>charset=windows-1252 in text</p>", "<p>charset=windows-1252 in text</p>")]
        public void The_meta_charset_is_rewritten_to_utf8(string html, string expected)
        {
            Parse(new MailContent(html, null, null)).Text.Should().Be(expected);
        }

        // ---- large parts --------------------------------------------------------

        [Fact]
        public void A_large_part_round_trips_without_excessive_allocation()
        {
            var data = new byte[5 * 1024 * 1024];
            new Random(42).NextBytes(data);
            var content = new MailContent("<p>x</p>", null, new[] { new MailPart(1, "big.bin", "application/octet-stream", data) });
            var builder = new MimeEntityBuilder();

            long before = GC.GetAllocatedBytesForCurrentThread();
            var bytes = builder.Build(content);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

            allocated.Should().BeLessThan((long)(bytes.Length * 2.5));
            MimeTestReader.Parse(bytes).Children[1].Body.Should().Equal(data);
        }

        [Fact]
        public void A_regular_part_with_a_content_id_stays_an_attachment_and_keeps_its_content_id()
        {
            var unflagged = new MailPart(2, "logo.png", "image/png", Png, isInline: false, contentId: "logo@x");
            var content = new MailContent("<html><body><img src=\"cid:logo@x\"></body></html>", null, new[] { unflagged });

            var root = MimeTestReader.Parse(new MimeEntityBuilder().Build(content));

            root.MediaType.Should().Be("multipart/mixed");
            root.Children[1].Disposition.Should().Be("attachment");
            root.Children[1].ContentId.Should().Be("<logo@x>");
        }

        // ---- model --------------------------------------------------------------

        [Fact]
        public void Mail_part_normalises_the_content_id_inline_or_not()
        {
            new MailPart(1, "a", null, null, isInline: true, contentId: "<cid@x>").ContentId.Should().Be("cid@x");
            new MailPart(1, "a", null, null, isInline: false, contentId: "<cid@x>").ContentId.Should().Be("cid@x");
            new MailPart(1, null, null, null).FileName.Should().BeEmpty();
            new MailPart(1, "a", null, null).Data.Should().BeEmpty();
        }

        [Fact]
        public void Mail_content_sums_part_bytes_and_tolerates_null_parts()
        {
            var content = new MailContent(null, null, new[] { Attachment(), null, InlineImage() });

            content.Parts.Should().HaveCount(2);
            content.TotalPartBytes.Should().Be(Pdf.Length + Png.Length);
            new MailContent(null, null, null).Parts.Should().BeEmpty();
        }
    }
}
