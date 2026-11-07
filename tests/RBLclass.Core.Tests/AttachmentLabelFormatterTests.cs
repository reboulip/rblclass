using System;
using System.Net;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class AttachmentLabelFormatterTests
    {
        private static readonly AttachmentLabelOptions Options = new AttachmentLabelOptions(
            "Former attachment:", "Former attachments:", "Saved to {0}", "Deleted on {0}", "yyyy-MM-dd");

        private static readonly DateTime When = new DateTime(2026, 6, 21, 9, 30, 0);

        private static MailItemRef Item() => new MailItemRef("s1", "e1", "subject");

        private static AttachmentDisposition Saved(string name, string dir) =>
            new AttachmentDisposition(Item(), 1, name, AttachmentDispositionAction.SaveTo, dir);

        private static AttachmentDisposition Deleted(string name) =>
            new AttachmentDisposition(Item(), 1, name, AttachmentDispositionAction.Delete);

        [Fact]
        public void Returns_null_for_an_empty_list()
        {
            AttachmentLabelFormatter.Format(new AttachmentDisposition[0], Options, When).Should().BeNull();
        }

        [Fact]
        public void Single_saved_attachment_uses_the_singular_header()
        {
            var html = AttachmentLabelFormatter.Format(
                new[] { Saved("report.pdf", @"C:\Docs") }, Options, When);

            html.Should().Contain("Former attachment:");
            html.Should().NotContain("Former attachments:");
            html.Should().Contain("report.pdf");
            html.Should().Contain("Saved to ");
            html.Should().Contain("href=\"file:///C:/Docs\"");
            html.Should().Contain(@">C:\Docs</a>");
        }

        [Fact]
        public void Multiple_attachments_use_the_plural_header()
        {
            var html = AttachmentLabelFormatter.Format(
                new[] { Saved("a.pdf", @"C:\X"), Deleted("b.pdf") }, Options, When);

            html.Should().Contain("Former attachments:");
            html.Should().Contain("href=\"file:///C:/X\"");
            html.Should().Contain(@">C:\X</a>");
            html.Should().Contain("Deleted on 2026-06-21");
        }

        [Fact]
        public void SaveTo_path_is_rendered_as_a_file_uri_hyperlink()
        {
            var html = AttachmentLabelFormatter.Format(
                new[] { Saved("invoice.pdf", @"C:\Archive\2026") }, Options, When);

            html.Should().Contain("href=\"file:///C:/Archive/2026\"");
            html.Should().Contain(@">C:\Archive\2026</a>");
            html.Should().NotContain(@"Saved to C:\Archive\2026");
        }

        // HtmlEncode turns Latin-1 characters into numeric entities, so assert on
        // the decoded markup: that is the href value a mail client hands the shell.
        private static string DecodedHtmlFor(string directory) =>
            WebUtility.HtmlDecode(AttachmentLabelFormatter.Format(
                new[] { Saved("file.pdf", directory) }, Options, When));

        [Fact]
        public void SaveTo_path_keeps_accented_characters_raw_in_the_href()
        {
            var html = DecodedHtmlFor(@"C:\Documents\Résumé");

            html.Should().Contain("href=\"file:///C:/Documents/Résumé\"");
            html.Should().Contain(@">C:\Documents\Résumé</a>");
            html.Should().NotContain("%C3%A9");
        }

        [Fact]
        public void SaveTo_path_percent_encodes_spaces_in_the_href_only()
        {
            var html = DecodedHtmlFor(@"C:\My Folder\archive");

            html.Should().Contain("href=\"file:///C:/My%20Folder/archive\"");
            html.Should().Contain(@">C:\My Folder\archive</a>");
        }

        [Fact]
        public void SaveTo_UNC_path_uses_the_server_as_the_uri_authority()
        {
            var html = DecodedHtmlFor(@"\\server\partage\Résumé");

            html.Should().Contain("href=\"file://server/partage/Résumé\"");
            html.Should().NotContain("%C3%A9");
        }

        [Fact]
        public void SaveTo_path_percent_encodes_uri_reserved_characters_once()
        {
            var html = DecodedHtmlFor(@"C:\100% #1?");

            html.Should().Contain("href=\"file:///C:/100%25%20%231%3F\"");
        }

        [Fact]
        public void SaveTo_path_html_encodes_the_href_attribute()
        {
            var html = AttachmentLabelFormatter.Format(
                new[] { Saved("file.pdf", @"C:\R&D") }, Options, When);

            html.Should().Contain("href=\"file:///C:/R&amp;D\"");
        }

        [Fact]
        public void Deleted_attachment_renders_the_date()
        {
            var html = AttachmentLabelFormatter.Format(new[] { Deleted("x.docx") }, Options, When);
            html.Should().Contain("Deleted on 2026-06-21");
        }

        [Fact]
        public void The_block_is_wrapped_in_a_div()
        {
            var html = AttachmentLabelFormatter.Format(new[] { Deleted("x.docx") }, Options, When);
            html.Should().StartWith("<div").And.EndWith("</div>");
        }
    }
}
