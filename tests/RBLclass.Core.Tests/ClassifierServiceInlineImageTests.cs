using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>Body images in the disposition flows (v2.11.0.0 C2).</summary>
    public class ClassifierServiceInlineImageTests
    {
        private const string Placeholder = "[Image removed]";
        private static readonly MailItemRef Mail = new MailItemRef("s1", "e1", "subject");
        private static readonly FolderNode Dest = new FolderNode("s1", "d1", null, "d1", "Archive / A", isLeaf: true);

        private static AttachmentDisposition Image(int id, string cid, AttachmentDispositionAction action, string dir = null) =>
            new AttachmentDisposition(Mail, id, "image" + id + ".png", action, dir, cid, Placeholder);

        private static AttachmentDisposition File(int id, AttachmentDispositionAction action, string dir = null) =>
            new AttachmentDisposition(Mail, id, "file" + id + ".pdf", action, dir);

        private static AttachmentLabelOptions Labels() =>
            new AttachmentLabelOptions("Former attachment:", "Former attachments:", "Saved to {0}", "Deleted on {0}", "yyyy-MM-dd");

        private static IMailStore Store()
        {
            var store = Substitute.For<IMailStore>();
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(new MailItemRef("s1", "filed", "subject"));
            store.CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(new MailItemRef("s1", "copy", "subject"));
            store.RemoveInlineImages(Arg.Any<MailItemRef>(), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>()).Returns(true);
            store.RemoveAttachment(Arg.Any<MailItemRef>(), Arg.Any<int>()).Returns(true);
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(true);
            store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), Arg.Any<int>(), Arg.Any<string>()).Returns(true);
            return store;
        }

        private static ClassifyResult Classify(IMailStore store, bool removeAttachments, bool keepCopy,
                                               AttachmentLabelOptions labels, params AttachmentDisposition[] rows) =>
            new ClassifierService(store).Classify(new ClassifyRequest(
                new[] { Mail }, new[] { Dest },
                new ClassifyOptions(keepCopy: keepCopy, removeAttachments: removeAttachments),
                attachmentDispositions: rows, labelOptions: labels));

        private static IReadOnlyList<string> Ids(params string[] ids) =>
            Arg.Is<IReadOnlyList<string>>(l => l.OrderBy(x => x).SequenceEqual(ids.OrderBy(x => x)));

        [Fact]
        public void A_deleted_body_image_is_removed_through_the_body_not_by_index()
        {
            var store = Store();

            Classify(store, true, false, null, Image(1, "img@x", AttachmentDispositionAction.Delete));

            store.Received(1).RemoveInlineImages(Arg.Is<MailItemRef>(r => r.EntryId == "filed"), Ids("img@x"), Placeholder);
            store.DidNotReceive().RemoveAttachment(Arg.Any<MailItemRef>(), Arg.Any<int>());
            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void A_kept_body_image_is_never_touched()
        {
            var store = Store();

            Classify(store, true, false, null, Image(1, "img@x", AttachmentDispositionAction.Keep));

            store.DidNotReceive().RemoveInlineImages(Arg.Any<MailItemRef>(), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>());
        }

        [Fact]
        public void Silent_removal_never_reaches_body_images()
        {
            var store = Store();

            Classify(store, true, false, null);

            store.Received(1).RemoveAttachments(Arg.Any<MailItemRef>());
            store.DidNotReceive().RemoveInlineImages(Arg.Any<MailItemRef>(), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>());
        }

        [Fact]
        public void Body_images_are_handled_even_when_attachment_removal_is_off()
        {
            var store = Store();

            Classify(store, false, false, null, Image(1, "img@x", AttachmentDispositionAction.Delete));

            store.Received(1).RemoveInlineImages(Arg.Any<MailItemRef>(), Ids("img@x"), Placeholder);
            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void With_keep_a_copy_only_the_filed_copy_loses_the_image()
        {
            var store = Store();

            Classify(store, true, true, null, Image(1, "img@x", AttachmentDispositionAction.Delete));

            store.Received(1).RemoveInlineImages(Arg.Is<MailItemRef>(r => r.EntryId == "copy"), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>());
            store.DidNotReceive().RemoveInlineImages(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>());
        }

        [Fact]
        public void Regular_attachments_are_removed_before_body_images()
        {
            var store = Store();

            Classify(store, true, false, null,
                File(1, AttachmentDispositionAction.Delete),
                File(2, AttachmentDispositionAction.Keep),
                Image(3, "img@x", AttachmentDispositionAction.Delete));

            Received.InOrder(() =>
            {
                store.RemoveAttachment(Arg.Any<MailItemRef>(), 1);
                store.RemoveInlineImages(Arg.Any<MailItemRef>(), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>());
            });
            store.DidNotReceive().RemoveAttachment(Arg.Any<MailItemRef>(), 3);
        }

        [Fact]
        public void Only_body_image_rows_never_trigger_strip_all()
        {
            var store = Store();

            Classify(store, true, false, null, Image(1, "img@x", AttachmentDispositionAction.Delete));

            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void A_failed_save_blocks_body_image_removal_too()
        {
            var store = Store();
            store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), 1, Arg.Any<string>()).Returns(false);

            Classify(store, true, false, null,
                File(1, AttachmentDispositionAction.SaveTo, @"C:\Out"),
                Image(2, "img@x", AttachmentDispositionAction.Delete));

            store.DidNotReceive().RemoveInlineImages(Arg.Any<MailItemRef>(), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>());
        }

        [Fact]
        public void A_saved_body_image_is_saved_then_removed()
        {
            var store = Store();

            Classify(store, true, false, null, Image(1, "img@x", AttachmentDispositionAction.SaveTo, @"C:\Out"));

            Received.InOrder(() =>
            {
                store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), 1, @"C:\Out");
                store.RemoveInlineImages(Arg.Any<MailItemRef>(), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>());
            });
        }

        [Fact]
        public void An_image_also_used_by_a_kept_row_is_not_removed()
        {
            var store = Store();

            Classify(store, true, false, null,
                Image(1, "shared@x", AttachmentDispositionAction.Delete),
                Image(2, "shared@x", AttachmentDispositionAction.Keep),
                Image(3, "gone@x", AttachmentDispositionAction.Delete));

            store.Received(1).RemoveInlineImages(Arg.Any<MailItemRef>(), Ids("gone@x"), Arg.Any<string>());
        }

        [Fact]
        public void A_failed_body_image_removal_is_left_out_of_the_label()
        {
            var store = Store();
            store.RemoveInlineImages(Arg.Any<MailItemRef>(), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>()).Returns(false);

            Classify(store, true, false, Labels(),
                File(1, AttachmentDispositionAction.Delete),
                File(2, AttachmentDispositionAction.Keep),
                Image(3, "img@x", AttachmentDispositionAction.Delete));

            store.Received(1).AppendHtmlNote(Arg.Any<MailItemRef>(),
                Arg.Is<string>(s => s.Contains("file1.pdf") && !s.Contains("image3.png")));
        }

        [Fact]
        public void A_removed_body_image_is_listed_in_the_label()
        {
            var store = Store();

            Classify(store, true, false, Labels(), Image(1, "img@x", AttachmentDispositionAction.Delete));

            store.Received(1).AppendHtmlNote(Arg.Any<MailItemRef>(), Arg.Is<string>(s => s.Contains("image1.png")));
        }

        [Fact]
        public void Manage_attachments_removes_a_deleted_body_image_in_place()
        {
            var store = Store();

            var result = new ClassifierService(store).ApplyAttachmentDispositions(
                new[] { Image(1, "img@x", AttachmentDispositionAction.Delete) });

            store.Received(1).RemoveInlineImages(Mail, Ids("img@x"), Placeholder);
            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
            result.Applied.Should().Be(1);
        }

        [Fact]
        public void Manage_attachments_leaves_a_saved_body_image_when_not_stripping_saved_ones()
        {
            var store = Store();

            new ClassifierService(store).ApplyAttachmentDispositions(
                new[] { Image(1, "img@x", AttachmentDispositionAction.SaveTo, @"C:\Out") },
                stripSavedAttachments: false);

            store.Received(1).SaveAttachmentToFile(Mail, 1, @"C:\Out");
            store.DidNotReceive().RemoveInlineImages(Arg.Any<MailItemRef>(), Arg.Any<IReadOnlyList<string>>(), Arg.Any<string>());
        }
    }
}
