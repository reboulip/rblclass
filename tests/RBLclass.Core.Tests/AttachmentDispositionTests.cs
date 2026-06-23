using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class AttachmentDispositionTests
    {
        private static MailItemRef Item(string entry) => new MailItemRef("s1", entry, entry);
        private static FolderNode Dest(string entry, string path) =>
            new FolderNode("s1", entry, null, entry, path, true);
        private static readonly FolderNode D1 = Dest("d1", "Archive");

        private static AttachmentDisposition Del(MailItemRef it, int id) =>
            new AttachmentDisposition(it, id, "f" + id + ".pdf", AttachmentDispositionAction.Delete);
        private static AttachmentDisposition Save(MailItemRef it, int id, string dir) =>
            new AttachmentDisposition(it, id, "f" + id + ".pdf", AttachmentDispositionAction.SaveTo, dir);

        private static IMailStore StoreWithWorkingMove()
        {
            var store = Substitute.For<IMailStore>();
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>())
                 .Returns(ci => new MailItemRef(((FolderNode)ci[1]).StoreId,
                     "moved-" + ((MailItemRef)ci[0]).EntryId, ((MailItemRef)ci[0]).Subject));
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(true);
            store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), Arg.Any<int>(), Arg.Any<string>()).Returns(true);
            return store;
        }

        [Fact]
        public void AttachmentInfo_is_not_inline_by_default()
        {
            new AttachmentInfo(1, "report.pdf", 100).IsInline.Should().BeFalse();
        }

        [Fact]
        public void AttachmentInfo_carries_the_inline_flag_when_set()
        {
            new AttachmentInfo(1, "logo.png", 50, isInline: true).IsInline.Should().BeTrue();
        }

        [Fact]
        public void Delete_disposition_strips_but_never_saves()
        {
            var store = StoreWithWorkingMove();
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true,
                attachmentDispositions: new[] { Del(Item("e1"), 1) }));

            store.DidNotReceive().SaveAttachmentToFile(Arg.Any<MailItemRef>(), Arg.Any<int>(), Arg.Any<string>());
            store.Received(1).RemoveAttachments(Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"));
        }

        [Fact]
        public void SaveTo_disposition_saves_then_strips_the_filed_item()
        {
            var store = StoreWithWorkingMove();
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true,
                attachmentDispositions: new[] { Save(Item("e1"), 1, @"C:\out") }));

            Received.InOrder(() =>
            {
                store.SaveAttachmentToFile(Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"), 1, @"C:\out");
                store.RemoveAttachments(Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"));
            });
        }

        [Fact]
        public void No_disposition_for_the_item_falls_back_to_stripping_all()
        {
            var store = StoreWithWorkingMove();
            var sut = new ClassifierService(store);

            // Disposition references a different item (e2); e1 has none.
            sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true,
                attachmentDispositions: new[] { Save(Item("e2"), 1, @"C:\out") }));

            store.DidNotReceive().SaveAttachmentToFile(
                Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"), Arg.Any<int>(), Arg.Any<string>());
            store.Received(1).RemoveAttachments(Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"));
        }

        [Fact]
        public void Encrypted_filed_item_is_reported_even_with_dispositions()
        {
            var store = StoreWithWorkingMove();
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(false); // encrypted: refused
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true,
                attachmentDispositions: new[] { Del(Item("e1"), 1) }));

            result.EncryptedStripSkips.Should().Be(1);
            result.ItemsProcessed.Should().Be(1);
        }

        [Fact]
        public void A_failed_save_never_strips_so_no_attachment_is_lost()
        {
            var store = StoreWithWorkingMove();
            store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), Arg.Any<int>(), Arg.Any<string>())
                 .Returns(false); // the save fails (e.g. bad path)
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true,
                attachmentDispositions: new[] { Save(Item("e1"), 1, @"C:\out") }));

            // Strip must NOT run - the attachment stays on the mail, nothing lost.
            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
            result.Errors.Should().Be(1);
        }

        [Fact]
        public void A_label_is_written_to_the_filed_item_after_disposition()
        {
            var store = StoreWithWorkingMove();
            var sut = new ClassifierService(store);
            var labelOptions = new AttachmentLabelOptions(
                "Former attachment:", "Former attachments:", "Saved to {0}", "Deleted on {0}", "yyyy-MM-dd");

            sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true,
                attachmentDispositions: new[] { Del(Item("e1"), 1) },
                labelOptions: labelOptions));

            store.Received(1).AppendHtmlNote(
                Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"),
                Arg.Is<string>(s => !string.IsNullOrEmpty(s)));
        }

        [Fact]
        public void No_label_is_written_when_the_strip_is_refused()
        {
            var store = StoreWithWorkingMove();
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(false); // encrypted
            var sut = new ClassifierService(store);
            var labelOptions = new AttachmentLabelOptions(
                "Former attachment:", "Former attachments:", "Saved to {0}", "Deleted on {0}", "yyyy-MM-dd");

            sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true,
                attachmentDispositions: new[] { Del(Item("e1"), 1) },
                labelOptions: labelOptions));

            store.DidNotReceive().AppendHtmlNote(Arg.Any<MailItemRef>(), Arg.Any<string>());
        }

        [Fact]
        public void SaveTo_acts_on_the_copy_when_keeping_a_copy()
        {
            var store = StoreWithWorkingMove();
            var copy = new MailItemRef("s1", "copy-e1", "e1");
            store.CopyItemToFolder(Arg.Any<MailItemRef>(), D1).Returns(copy);
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new[] { D1 },
                keepCopy: true, removeAttachments: true,
                attachmentDispositions: new[] { Save(Item("e1"), 1, @"C:\out") }));

            // The save and strip act on the filed copy, never the kept original.
            store.Received(1).SaveAttachmentToFile(
                Arg.Is<MailItemRef>(m => m.EntryId == "copy-e1"), 1, @"C:\out");
            store.DidNotReceive().SaveAttachmentToFile(
                Arg.Is<MailItemRef>(m => m.EntryId == "e1"), Arg.Any<int>(), Arg.Any<string>());
        }
    }
}
