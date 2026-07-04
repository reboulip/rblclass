using System.Collections.Generic;
using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// v2.8.0.0 A2: standalone attachment-disposition apply (no classify move).
    /// Mirrors the classify-path attachment rules but operates on the item itself.
    /// </summary>
    public class ClassifierServiceApplyAttachmentDispositionsTests
    {
        private static MailItemRef Item(string entry) => new MailItemRef("s1", entry, entry);

        private static AttachmentDisposition Disp(MailItemRef item, int id,
            AttachmentDispositionAction action, string dir = null) =>
            new AttachmentDisposition(item, id, "file" + id + ".pdf", action, dir);

        private static AttachmentLabelOptions MinimalLabels() =>
            new AttachmentLabelOptions("Former attachment", "Former attachments",
                "Saved to {0}", "Deleted on {0}", "yyyy-MM-dd");

        [Fact]
        public void AllDelete_calls_RemoveAttachments_and_returns_applied_one()
        {
            var store = Substitute.For<IMailStore>();
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(true);
            var sut = new ClassifierService(store);
            var e1 = Item("e1");

            var result = sut.ApplyAttachmentDispositions(new[]
            {
                Disp(e1, 1, AttachmentDispositionAction.Delete),
                Disp(e1, 2, AttachmentDispositionAction.Delete),
            });

            store.Received(1).RemoveAttachments(Arg.Is<MailItemRef>(m => m.EntryId == "e1"));
            result.Applied.Should().Be(1);
            result.Errors.Should().Be(0);
            result.EncryptedSkips.Should().Be(0);
        }

        [Fact]
        public void AllSaveTo_saves_first_then_strips()
        {
            var store = Substitute.For<IMailStore>();
            store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), Arg.Any<int>(), Arg.Any<string>())
                 .Returns(true);
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(true);
            var sut = new ClassifierService(store);
            var e1 = Item("e1");

            var result = sut.ApplyAttachmentDispositions(new[]
            {
                Disp(e1, 1, AttachmentDispositionAction.SaveTo, @"C:\dst"),
            });

            Received.InOrder(() =>
            {
                store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), 1, @"C:\dst");
                store.RemoveAttachments(Arg.Any<MailItemRef>());
            });
            result.Applied.Should().Be(1);
        }

        [Fact]
        public void WithKeep_removes_only_nonKeep()
        {
            var store = Substitute.For<IMailStore>();
            store.RemoveAttachment(Arg.Any<MailItemRef>(), Arg.Any<int>()).Returns(true);
            var sut = new ClassifierService(store);
            var e1 = Item("e1");

            var result = sut.ApplyAttachmentDispositions(new[]
            {
                Disp(e1, 2, AttachmentDispositionAction.Keep),
                Disp(e1, 1, AttachmentDispositionAction.Delete),
            });

            store.Received(1).RemoveAttachment(Arg.Any<MailItemRef>(), 1);
            store.DidNotReceive().RemoveAttachment(Arg.Any<MailItemRef>(), 2);
            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
            result.Applied.Should().Be(1);
        }

        [Fact]
        public void WithKeep_multiple_deletes_removes_highest_id_first()
        {
            var store = Substitute.For<IMailStore>();
            store.RemoveAttachment(Arg.Any<MailItemRef>(), Arg.Any<int>()).Returns(true);
            var sut = new ClassifierService(store);
            var e1 = Item("e1");

            sut.ApplyAttachmentDispositions(new[]
            {
                Disp(e1, 3, AttachmentDispositionAction.Delete),
                Disp(e1, 1, AttachmentDispositionAction.Delete),
                Disp(e1, 2, AttachmentDispositionAction.Keep),
            });

            Received.InOrder(() =>
            {
                store.RemoveAttachment(Arg.Any<MailItemRef>(), 3);
                store.RemoveAttachment(Arg.Any<MailItemRef>(), 1);
            });
        }

        [Fact]
        public void SaveFails_blocks_strip_and_counts_error()
        {
            var store = Substitute.For<IMailStore>();
            store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), Arg.Any<int>(), Arg.Any<string>())
                 .Returns(false);
            var sut = new ClassifierService(store);
            var e1 = Item("e1");

            var result = sut.ApplyAttachmentDispositions(new[]
            {
                Disp(e1, 1, AttachmentDispositionAction.SaveTo, @"C:\dst"),
            });

            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
            store.DidNotReceive().RemoveAttachment(Arg.Any<MailItemRef>(), Arg.Any<int>());
            result.Applied.Should().Be(0);
            result.Errors.Should().Be(1);
        }

        [Fact]
        public void EncryptedItem_counts_encrypted_skip()
        {
            var store = Substitute.For<IMailStore>();
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(false); // S/MIME guard
            var sut = new ClassifierService(store);
            var e1 = Item("e1");

            var result = sut.ApplyAttachmentDispositions(new[]
            {
                Disp(e1, 1, AttachmentDispositionAction.Delete),
            });

            result.Applied.Should().Be(0);
            result.EncryptedSkips.Should().Be(1);
            result.Errors.Should().Be(0);
        }

        [Fact]
        public void LabelOptions_appends_note_after_strip()
        {
            var store = Substitute.For<IMailStore>();
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(true);
            var sut = new ClassifierService(store);
            var e1 = Item("e1");

            sut.ApplyAttachmentDispositions(new[]
            {
                Disp(e1, 1, AttachmentDispositionAction.Delete),
            }, MinimalLabels());

            store.Received(1).AppendHtmlNote(Arg.Is<MailItemRef>(m => m.EntryId == "e1"),
                Arg.Is<string>(s => !string.IsNullOrEmpty(s)));
        }

        [Fact]
        public void MultipleItems_each_processed_independently()
        {
            var store = Substitute.For<IMailStore>();
            var e1 = Item("e1");
            var e2 = Item("e2");
            // e1 save fails -> error, strip blocked; e2 delete succeeds.
            store.SaveAttachmentToFile(Arg.Is<MailItemRef>(m => m.EntryId == "e1"),
                Arg.Any<int>(), Arg.Any<string>()).Returns(false);
            store.RemoveAttachments(Arg.Is<MailItemRef>(m => m.EntryId == "e2")).Returns(true);
            var sut = new ClassifierService(store);

            var result = sut.ApplyAttachmentDispositions(new[]
            {
                Disp(e1, 1, AttachmentDispositionAction.SaveTo, @"C:\dst"),
                Disp(e2, 1, AttachmentDispositionAction.Delete),
            });

            result.Applied.Should().Be(1);
            result.Errors.Should().Be(1);
        }

        [Fact]
        public void EmptyDispositions_returns_zero_result()
        {
            var store = Substitute.For<IMailStore>();
            var sut = new ClassifierService(store);

            var result = sut.ApplyAttachmentDispositions(null);

            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
            result.Applied.Should().Be(0);
            result.Errors.Should().Be(0);
            result.EncryptedSkips.Should().Be(0);
        }
    }
}
