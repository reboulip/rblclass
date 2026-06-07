using System;
using System.Collections.Generic;
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
