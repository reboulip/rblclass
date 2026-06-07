using System;
using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class ClassifierServiceTests
    {
        private static MailItemRef Item(string entry) => new MailItemRef("s1", entry, entry);

        private static FolderNode Dest(string entry, string path) =>
            new FolderNode("s1", entry, null, entry, path, isLeaf: true);

        private static readonly FolderNode D1 = Dest("d1", "Archive / A");
        private static readonly FolderNode D2 = Dest("d2", "Archive / B");

        [Fact]
        public void Copies_each_item_to_each_destination_and_deletes_originals()
        {
            var store = Substitute.For<IMailStore>();
            var sut = new ClassifierService(store);

            var request = new ClassifyRequest(
                new[] { Item("e1"), Item("e2") },
                new[] { D1, D2 },
                keepCopy: false,
                removeAttachments: false);

            var result = sut.Classify(request);

            // 2 items x 2 destinations = 4 copies.
            store.Received(4).CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>());
            store.Received(1).CopyItemToFolder(Arg.Is<MailItemRef>(m => m.EntryId == "e1"), D1);
            store.Received(1).CopyItemToFolder(Arg.Is<MailItemRef>(m => m.EntryId == "e2"), D2);
            store.Received(2).DeleteItem(Arg.Any<MailItemRef>());
            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());

            result.ItemsProcessed.Should().Be(2);
            result.CopiesMade.Should().Be(4);
            result.OriginalsDeleted.Should().Be(2);
            result.Errors.Should().Be(0);
        }

        [Fact]
        public void Keep_copy_does_not_delete_originals()
        {
            var store = Substitute.For<IMailStore>();
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: true, removeAttachments: false));

            store.Received(1).CopyItemToFolder(Arg.Any<MailItemRef>(), D1);
            store.DidNotReceive().DeleteItem(Arg.Any<MailItemRef>());
            result.OriginalsDeleted.Should().Be(0);
        }

        [Fact]
        public void Remove_attachments_strips_the_filed_copy_not_the_original()
        {
            var store = Substitute.For<IMailStore>();
            var item = Item("e1");
            var filedCopy = new MailItemRef("s1", "copy-in-dest", "e1");
            store.CopyItemToFolder(item, D1).Returns(filedCopy);
            var sut = new ClassifierService(store);

            // keep a copy + remove attachments: original must keep its attachments.
            sut.Classify(new ClassifyRequest(
                new[] { item }, new[] { D1 },
                keepCopy: true, removeAttachments: true));

            store.Received(1).RemoveAttachments(filedCopy);     // the copy is stripped
            store.DidNotReceive().RemoveAttachments(item);      // the original is NOT
            store.DidNotReceive().DeleteItem(item);             // and is kept

            Received.InOrder(() =>
            {
                store.CopyItemToFolder(item, D1);
                store.RemoveAttachments(filedCopy);
            });
        }

        [Fact]
        public void Remove_attachments_off_strips_nothing()
        {
            var store = Substitute.For<IMailStore>();
            store.CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>())
                 .Returns(new MailItemRef("s1", "copy", "x"));
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: false));

            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void Preflight_without_widening_just_dedupes_the_selection()
        {
            var store = Substitute.For<IMailStore>();
            var sut = new ClassifierService(store);

            var preflight = sut.Preflight(new[] { Item("e1"), Item("e1"), Item("e2") }, widenConversation: false);

            preflight.Items.Select(i => i.EntryId).Should().BeEquivalentTo(new[] { "e1", "e2" });
            store.DidNotReceive().GetConversationSiblings(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void Preflight_with_widening_adds_conversation_siblings_deduped()
        {
            var store = Substitute.For<IMailStore>();
            var e1 = Item("e1");
            var e2 = Item("e2");
            var sibling = new MailItemRef("s1", "sib", "reply");
            // e1's siblings include e2 (already selected) and a genuinely new one;
            // e2 reports the same new sibling back (e.g. both ends of a thread).
            store.GetConversationSiblings(e1).Returns(new[] { e2, sibling });
            store.GetConversationSiblings(e2).Returns(new[] { sibling });
            var sut = new ClassifierService(store);

            var preflight = sut.Preflight(new[] { e1, e2 }, widenConversation: true);

            preflight.Items.Select(i => i.EntryId).Should().BeEquivalentTo(new[] { "e1", "e2", "sib" });
            preflight.Items.Should().HaveCount(3); // no duplicates despite the cross-reported sibling
        }

        [Fact]
        public void Preflight_reports_the_flagged_incomplete_subset()
        {
            var store = Substitute.For<IMailStore>();
            var flagged = Item("e1");
            var done = Item("e2");
            var unflagged = Item("e3");
            store.IsFlaggedIncomplete(flagged).Returns(true);
            store.IsFlaggedIncomplete(done).Returns(false);
            store.IsFlaggedIncomplete(unflagged).Returns(false);
            var sut = new ClassifierService(store);

            var preflight = sut.Preflight(new[] { flagged, done, unflagged }, widenConversation: false);

            preflight.FlaggedIncomplete.Should().BeEquivalentTo(new[] { flagged });
        }

        [Fact]
        public void Classify_marks_the_filed_copy_complete_for_flagged_items_only()
        {
            var store = Substitute.For<IMailStore>();
            var flagged = Item("e1");
            var plain = Item("e2");
            var flaggedCopy = new MailItemRef("s1", "copy-flagged", "e1");
            var plainCopy = new MailItemRef("s1", "copy-plain", "e2");
            store.IsFlaggedIncomplete(flagged).Returns(true);
            store.IsFlaggedIncomplete(plain).Returns(false);
            store.CopyItemToFolder(flagged, D1).Returns(flaggedCopy);
            store.CopyItemToFolder(plain, D1).Returns(plainCopy);
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(
                new[] { flagged, plain }, new[] { D1 },
                keepCopy: false, removeAttachments: false, markTasksComplete: true));

            store.Received(1).MarkTaskComplete(flaggedCopy);
            store.DidNotReceive().MarkTaskComplete(plainCopy);
            store.DidNotReceive().MarkTaskComplete(flagged); // never the original
        }

        [Fact]
        public void Classify_marks_nothing_complete_when_not_requested()
        {
            var store = Substitute.For<IMailStore>();
            var flagged = Item("e1");
            var copy = new MailItemRef("s1", "copy", "e1");
            store.IsFlaggedIncomplete(flagged).Returns(true);
            store.CopyItemToFolder(flagged, D1).Returns(copy);
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(
                new[] { flagged }, new[] { D1 },
                keepCopy: false, removeAttachments: false, markTasksComplete: false));

            store.DidNotReceive().MarkTaskComplete(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void Empty_items_or_destinations_is_a_no_op()
        {
            var store = Substitute.For<IMailStore>();
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(new MailItemRef[0], new[] { D1 }, false, false))
               .CopiesMade.Should().Be(0);
            sut.Classify(new ClassifyRequest(new[] { Item("e1") }, new FolderNode[0], false, false))
               .CopiesMade.Should().Be(0);

            store.DidNotReceive().CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>());
            store.DidNotReceive().DeleteItem(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void One_failing_item_is_counted_and_the_rest_still_process()
        {
            var store = Substitute.For<IMailStore>();
            // Fail the copy for e1 only.
            store.When(s => s.CopyItemToFolder(Arg.Is<MailItemRef>(m => m.EntryId == "e1"), Arg.Any<FolderNode>()))
                 .Do(_ => throw new InvalidOperationException("boom"));
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1"), Item("e2") }, new[] { D1 },
                keepCopy: false, removeAttachments: false));

            result.Errors.Should().Be(1);
            result.ItemsProcessed.Should().Be(1); // e2
            result.OriginalsDeleted.Should().Be(1); // only e2's original
            store.Received(1).DeleteItem(Arg.Is<MailItemRef>(m => m.EntryId == "e2"));
            store.DidNotReceive().DeleteItem(Arg.Is<MailItemRef>(m => m.EntryId == "e1"));
        }
    }
}
