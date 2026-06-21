using System;
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
            html.Should().Contain(@"Saved to C:\Docs");
        }

        [Fact]
        public void Multiple_attachments_use_the_plural_header()
        {
            var html = AttachmentLabelFormatter.Format(
                new[] { Saved("a.pdf", @"C:\X"), Deleted("b.pdf") }, Options, When);

            html.Should().Contain("Former attachments:");
            html.Should().Contain(@"Saved to C:\X");
            html.Should().Contain("Deleted on 2026-06-21");
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
