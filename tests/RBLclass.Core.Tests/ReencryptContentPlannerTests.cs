using System;
using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class ReencryptContentPlannerTests
    {
        private static readonly MailItemRef Item = new MailItemRef("s1", "e1", "subject");
        private static readonly DateTime When = new DateTime(2026, 11, 14, 10, 0, 0);
        private static readonly AttachmentLabelOptions Label = new AttachmentLabelOptions(
            "Former attachment:", "Former attachments:", "Saved to {0}", "Deleted on {0}", "yyyy-MM-dd");

        private const string Banner =
            "<table border=\"0\" cellpadding=\"0\" width=\"100%\" style=\"background:#FFFFCC\">" +
            "<tr><td><b>CAUTION:</b> This email originated from outside the organisation. " +
            "Do not click links or open attachments unless you recognise the sender.</td></tr>" +
            "</table>";

        private static readonly MailPart Report = new MailPart(1, "rapport.txt", "text/plain", new byte[] { 1 });
        private static readonly MailPart Data = new MailPart(2, "donnees.csv", "text/csv", new byte[] { 2 });
        private static readonly MailPart Logo = new MailPart(3, "logo.png", "image/png", new byte[] { 3 }, isInline: true, contentId: "logo@x");

        private const string Html = "<html><body lang=FR><p>Bonjour</p><img src=\"cid:logo@x\"></body></html>";

        private static MailContent Content(string html = Html, string plain = null) =>
            new MailContent(html, plain, new[] { Report, Data, Logo });

        private static AttachmentDisposition Row(MailPart part, AttachmentDispositionAction action, string dir = null) =>
            new AttachmentDisposition(Item, part.AttachmentId, part.FileName, action, dir);

        private static AttachmentDisposition ImageRow(AttachmentDispositionAction action, string cid = "logo@x") =>
            new AttachmentDisposition(Item, Logo.AttachmentId, Logo.FileName, action, null,
                                      contentId: cid, inlinePlaceholder: "[Image supprimée]");

        private static ReencryptContentPlan Plan(MailContent content, bool remove, params AttachmentDisposition[] rows) =>
            ReencryptContentPlanner.Plan(content, remove, rows, null, Label, When);

        // ---- strip / keep -------------------------------------------------------

        [Fact]
        public void Strip_all_removes_every_non_inline_part_and_keeps_inline_ones()
        {
            var plan = Plan(Content(), true);

            plan.Content.Parts.Should().Equal(Logo);
            plan.RemovedPartCount.Should().Be(2);
            plan.LabelledRows.Should().BeEmpty();          // no rows: silent strip, no label
            plan.Content.HtmlBody.Should().Be(Html);
        }

        [Fact]
        public void Delete_rows_strip_all_and_are_labelled()
        {
            var rows = new[] { Row(Report, AttachmentDispositionAction.Delete), Row(Data, AttachmentDispositionAction.Delete) };

            var plan = Plan(Content(), true, rows);

            plan.Content.Parts.Should().Equal(Logo);
            plan.LabelledRows.Should().Equal(rows);
            plan.Content.HtmlBody.Should().StartWith("<html><body lang=FR><div").And.Contain("rapport.txt").And.Contain("donnees.csv");
        }

        [Fact]
        public void A_keep_row_switches_to_per_id_removal()
        {
            var keep = Row(Report, AttachmentDispositionAction.Keep);
            var delete = Row(Data, AttachmentDispositionAction.Delete);

            var plan = Plan(Content(), true, keep, delete);

            plan.Content.Parts.Should().Equal(Report, Logo);
            plan.LabelledRows.Should().Equal(delete);
            plan.AttachmentMapMismatch.Should().BeFalse();
        }

        [Fact]
        public void Save_to_rows_are_removed_and_labelled_with_their_folder()
        {
            var save = Row(Report, AttachmentDispositionAction.SaveTo, @"C:\Archive");

            var plan = Plan(Content(), true, save, Row(Data, AttachmentDispositionAction.Keep));

            plan.Content.Parts.Should().Equal(Data, Logo);
            plan.Content.HtmlBody.Should().Contain("Saved to").And.Contain("Archive");
        }

        [Fact]
        public void Remove_attachments_off_leaves_regular_parts_alone()
        {
            var plan = Plan(Content(), false, Row(Report, AttachmentDispositionAction.Delete));

            plan.Content.Parts.Should().Equal(Report, Data, Logo);
            plan.RemovedPartCount.Should().Be(0);
            plan.LabelledRows.Should().BeEmpty();
        }

        [Fact]
        public void Nothing_requested_changes_nothing()
        {
            var plan = Plan(Content(), false);

            plan.Content.Parts.Should().Equal(Report, Data, Logo);
            plan.Content.HtmlBody.Should().Be(Html);
        }

        // ---- mapping guard ------------------------------------------------------

        [Fact]
        public void A_missing_part_id_is_a_mapping_mismatch()
        {
            var ghost = new AttachmentDisposition(Item, 9, "ghost.txt", AttachmentDispositionAction.Delete);

            var plan = Plan(Content(), true, Row(Report, AttachmentDispositionAction.Keep), ghost);

            plan.AttachmentMapMismatch.Should().BeTrue();
        }

        [Fact]
        public void A_renamed_part_is_a_mapping_mismatch_and_is_not_removed()
        {
            var wrongName = new AttachmentDisposition(Item, Data.AttachmentId, "other.csv", AttachmentDispositionAction.Delete);

            var plan = Plan(Content(), true, Row(Report, AttachmentDispositionAction.Keep), wrongName);

            plan.AttachmentMapMismatch.Should().BeTrue();
            plan.Content.Parts.Should().Contain(Data);
        }

        // ---- body images --------------------------------------------------------

        [Fact]
        public void A_removed_body_image_is_replaced_dropped_and_labelled()
        {
            var image = ImageRow(AttachmentDispositionAction.Delete);

            var plan = Plan(Content(), false, image);

            plan.Content.Parts.Should().Equal(Report, Data);
            plan.Content.HtmlBody.Should().Contain("[Image supprim&#233;e]").And.NotContain("cid:logo@x");
            plan.LabelledRows.Should().Equal(image);
        }

        [Fact]
        public void A_kept_body_image_stays()
        {
            var plan = Plan(Content(), false, ImageRow(AttachmentDispositionAction.Keep));

            plan.Content.Parts.Should().Contain(Logo);
            plan.Content.HtmlBody.Should().Be(Html);
        }

        [Fact]
        public void A_body_image_not_referenced_in_the_html_is_kept()
        {
            var plan = Plan(Content(html: "<html><body><p>no image</p></body></html>"), false,
                            ImageRow(AttachmentDispositionAction.Delete));

            plan.Content.Parts.Should().Contain(Logo);
            plan.LabelledRows.Should().BeEmpty();
        }

        [Fact]
        public void Body_images_are_ignored_without_an_html_body()
        {
            var plan = Plan(Content(html: null, plain: "text"), false, ImageRow(AttachmentDispositionAction.Delete));

            plan.Content.Parts.Should().Contain(Logo);
            plan.Content.HtmlBody.Should().BeNull();
        }

        // ---- label placement ----------------------------------------------------

        [Fact]
        public void The_label_is_prepended_when_there_is_no_body_tag()
        {
            var plan = Plan(Content(html: "<p>no body tag</p>"), true, Row(Report, AttachmentDispositionAction.Delete));

            plan.Content.HtmlBody.Should().StartWith("<div").And.EndWith("<p>no body tag</p>");
        }

        [Fact]
        public void A_plain_mail_gets_an_html_body_carrying_the_label_and_keeps_its_plain_body()
        {
            var plan = Plan(Content(html: null, plain: "Line <1>\r\nLine 2"), true, Row(Report, AttachmentDispositionAction.Delete));

            plan.Content.PlainBody.Should().Be("Line <1>\r\nLine 2");
            plan.Content.HtmlBody.Should().StartWith("<html><body><div")
                .And.Contain("rapport.txt")
                .And.Contain("Line &lt;1&gt;");
        }

        [Fact]
        public void No_label_options_means_no_label()
        {
            var plan = ReencryptContentPlanner.Plan(Content(), true,
                new[] { Row(Report, AttachmentDispositionAction.Delete) }, null, null, When);

            plan.Content.HtmlBody.Should().Be(Html);
            plan.LabelledRows.Should().HaveCount(1);
        }

        // ---- banner -------------------------------------------------------------

        [Fact]
        public void The_banner_is_stripped_after_the_label()
        {
            string html = "<html><head></head><body>" + Banner + "<p>Real content stays.</p></body></html>";

            var plan = ReencryptContentPlanner.Plan(Content(html: html), true,
                new[] { Row(Report, AttachmentDispositionAction.Delete) }, Banner, Label, When);

            plan.BannerStripped.Should().BeTrue();
            plan.Content.HtmlBody.Should().NotContain("CAUTION").And.Contain("Real content stays.").And.Contain("rapport.txt");
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void No_banner_signature_strips_nothing(string signature)
        {
            var plan = ReencryptContentPlanner.Plan(Content(), false, null, signature, Label, When);

            plan.BannerStripped.Should().BeFalse();
            plan.Content.HtmlBody.Should().Be(Html);
        }

        [Fact]
        public void An_absent_banner_strips_nothing()
        {
            var plan = ReencryptContentPlanner.Plan(Content(), false, null, Banner, Label, When);

            plan.BannerStripped.Should().BeFalse();
        }

        // ---- purity -------------------------------------------------------------

        [Fact]
        public void The_input_is_never_changed_and_surviving_parts_are_reused()
        {
            var content = Content();

            var plan = Plan(content, true, Row(Report, AttachmentDispositionAction.Delete),
                            Row(Data, AttachmentDispositionAction.Keep), ImageRow(AttachmentDispositionAction.Delete));

            content.Parts.Should().Equal(Report, Data, Logo);
            content.HtmlBody.Should().Be(Html);
            plan.Content.Parts.Should().OnlyContain(p => ReferenceEquals(p, Data));
        }

        [Fact]
        public void A_null_content_throws()
        {
            Action plan = () => ReencryptContentPlanner.Plan(null, true, null, null, null, When);

            plan.Should().Throw<ArgumentNullException>();
        }
    }
}
