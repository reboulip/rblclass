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

        /// <summary>A store whose MoveItemToFolder returns a ref at the destination (the adapter contract).</summary>
        private static IMailStore StoreWithWorkingMove()
        {
            var store = Substitute.For<IMailStore>();
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>())
                 .Returns(ci => new MailItemRef(
                     ((FolderNode)ci[1]).StoreId,
                     "moved-" + ((MailItemRef)ci[0]).EntryId,
                     ((MailItemRef)ci[0]).Subject));
            return store;
        }

        [Fact]
        public void Without_keep_copy_the_original_moves_to_the_last_destination_and_extras_get_copies()
        {
            var store = StoreWithWorkingMove();
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1"), Item("e2") },
                new[] { D1, D2 },
                keepCopy: false,
                removeAttachments: false));

            // Per item: one copy to D1, then the original MOVES to D2 -
            // nothing is deleted (v2.2 move-based classify).
            store.Received(2).CopyItemToFolder(Arg.Any<MailItemRef>(), D1);
            store.Received(2).MoveItemToFolder(Arg.Any<MailItemRef>(), D2);
            store.DidNotReceive().CopyItemToFolder(Arg.Any<MailItemRef>(), D2);
            store.DidNotReceive().DeleteItem(Arg.Any<MailItemRef>());
            store.DidNotReceive().RemoveAttachments(Arg.Any<MailItemRef>());

            result.ItemsProcessed.Should().Be(2);
            result.CopiesMade.Should().Be(2);
            result.OriginalsMoved.Should().Be(2);
            result.Errors.Should().Be(0);
        }

        [Fact]
        public void Single_destination_without_keep_copy_is_a_pure_move()
        {
            var store = StoreWithWorkingMove();
            var sut = new ClassifierService(store);

            var item = Item("e1");
            var result = sut.Classify(new ClassifyRequest(
                new[] { item }, new[] { D1 },
                keepCopy: false, removeAttachments: false));

            // No transient copy, no delete - the Stormshield fix.
            store.Received(1).MoveItemToFolder(item, D1);
            store.DidNotReceive().CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>());
            store.DidNotReceive().DeleteItem(Arg.Any<MailItemRef>());

            result.ItemsProcessed.Should().Be(1);
            result.CopiesMade.Should().Be(0);
            result.OriginalsMoved.Should().Be(1);
        }

        [Fact]
        public void Copies_are_made_before_the_original_moves()
        {
            var store = StoreWithWorkingMove();
            var sut = new ClassifierService(store);
            var item = Item("e1");

            sut.Classify(new ClassifyRequest(
                new[] { item }, new[] { D1, D2 },
                keepCopy: false, removeAttachments: false));

            // The original is the copy source - it must still exist (i.e. not
            // yet have moved) when the copy is taken.
            Received.InOrder(() =>
            {
                store.CopyItemToFolder(item, D1);
                store.MoveItemToFolder(item, D2);
            });
        }

        [Fact]
        public void Keep_copy_copies_everywhere_and_never_moves_or_deletes_the_original()
        {
            var store = Substitute.For<IMailStore>();
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1, D2 },
                keepCopy: true, removeAttachments: false));

            store.Received(1).CopyItemToFolder(Arg.Any<MailItemRef>(), D1);
            store.Received(1).CopyItemToFolder(Arg.Any<MailItemRef>(), D2);
            store.DidNotReceive().MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>());
            store.DidNotReceive().DeleteItem(Arg.Any<MailItemRef>());
            result.OriginalsMoved.Should().Be(0);
            result.CopiesMade.Should().Be(2);
        }

        [Fact]
        public void Safety_copy_puts_a_copy_of_the_moved_item_in_the_source_stores_deleted_items()
        {
            var store = StoreWithWorkingMove();
            var deletedItems = Dest("del", "Deleted Items");
            store.GetDeletedItemsFolder("s1").Returns(deletedItems);
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(true);
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true,
                markTasksComplete: false, safetyCopy: true));

            // Copied from the moved item (at its destination) into Deleted
            // Items, and BEFORE stripping - the guardrail keeps attachments.
            Received.InOrder(() =>
            {
                store.MoveItemToFolder(Arg.Is<MailItemRef>(m => m.EntryId == "e1"), D1);
                store.CopyItemToFolder(Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"), deletedItems);
                store.RemoveAttachments(Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"));
            });
        }

        [Fact]
        public void Safety_copy_off_or_keep_copy_on_never_touches_deleted_items()
        {
            var store = StoreWithWorkingMove();
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: false)); // safetyCopy default off
            sut.Classify(new ClassifyRequest(
                new[] { Item("e2") }, new[] { D1 },
                keepCopy: true, removeAttachments: false,
                markTasksComplete: false, safetyCopy: true)); // keep-copy wins: no move, no guardrail

            store.DidNotReceive().GetDeletedItemsFolder(Arg.Any<string>());
        }

        [Fact]
        public void A_failed_safety_copy_does_not_fail_the_filing()
        {
            var store = StoreWithWorkingMove();
            store.GetDeletedItemsFolder("s1").Returns(Dest("del", "Deleted Items"));
            store.When(s => s.CopyItemToFolder(
                    Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"), Arg.Any<FolderNode>()))
                 .Do(_ => throw new InvalidOperationException("boom"));
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: false,
                markTasksComplete: false, safetyCopy: true));

            result.ItemsProcessed.Should().Be(1);
            result.OriginalsMoved.Should().Be(1);
            result.Errors.Should().Be(0); // the guardrail is best-effort
        }

        [Fact]
        public void Classify_records_an_undo_plan_with_moves_copies_and_flags()
        {
            var store = StoreWithWorkingMove();
            var source = Dest("src", "Inbox");
            store.GetParentFolder(Arg.Any<MailItemRef>()).Returns(source);
            store.IsFlaggedIncomplete(Arg.Any<MailItemRef>()).Returns(true);
            store.CopyItemToFolder(Arg.Any<MailItemRef>(), D1)
                 .Returns(new MailItemRef("s1", "copy-d1", "e1"));
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1, D2 },
                keepCopy: false, removeAttachments: false, markTasksComplete: true));

            var plan = result.Undo;
            plan.Should().NotBeNull();
            plan.Moves.Should().ContainSingle();
            plan.Moves[0].Current.EntryId.Should().Be("moved-e1");
            plan.Moves[0].SourceFolder.Should().BeSameAs(source);
            plan.CreatedCopies.Should().ContainSingle().Which.EntryId.Should().Be("copy-d1");
            plan.CompletedFlags.Select(f => f.EntryId)
                .Should().BeEquivalentTo(new[] { "copy-d1", "moved-e1" });
            plan.AttachmentStrips.Should().Be(0);
        }

        [Fact]
        public void Undo_restores_flags_then_deletes_copies_then_moves_items_back()
        {
            var store = StoreWithWorkingMove();
            var source = Dest("src", "Inbox");
            var movedRef = new MailItemRef("s1", "moved-e1", "e1");
            var copyRef = new MailItemRef("s1", "copy-d1", "e1");
            var sut = new ClassifierService(store);

            var plan = new ClassifyUndoPlan(
                new[] { new UndoableMove(movedRef, source) },
                new[] { copyRef },
                new[] { movedRef },
                attachmentStrips: 0);

            var result = sut.Undo(plan);

            Received.InOrder(() =>
            {
                store.MarkTaskIncomplete(movedRef);      // flags while refs are valid
                store.DeleteItemPermanently(copyRef);    // our duplicates go for good
                store.MoveItemToFolder(movedRef, source); // then the item goes home
            });
            result.MovesRestored.Should().Be(1);
            result.CopiesDeleted.Should().Be(1);
            result.FlagsRestored.Should().Be(1);
            result.Errors.Should().Be(0);
        }

        [Fact]
        public void Undo_counts_failures_but_keeps_going()
        {
            var store = StoreWithWorkingMove();
            var source = Dest("src", "Inbox");
            var copy1 = new MailItemRef("s1", "c1", "x");
            var copy2 = new MailItemRef("s1", "c2", "x");
            store.When(s => s.DeleteItemPermanently(copy1))
                 .Do(_ => throw new InvalidOperationException("boom"));
            var sut = new ClassifierService(store);

            var result = sut.Undo(new ClassifyUndoPlan(
                new[] { new UndoableMove(new MailItemRef("s1", "m1", "x"), source) },
                new[] { copy1, copy2 },
                new MailItemRef[0],
                attachmentStrips: 0));

            result.CopiesDeleted.Should().Be(1);  // copy2 still removed
            result.MovesRestored.Should().Be(1);  // the move still undone
            result.Errors.Should().Be(1);         // copy1's failure counted
        }

        [Fact]
        public void Keep_copy_classify_yields_a_copies_only_undo_plan()
        {
            var store = Substitute.For<IMailStore>();
            var copyRef = new MailItemRef("s1", "copy-d1", "e1");
            store.CopyItemToFolder(Arg.Any<MailItemRef>(), D1).Returns(copyRef);
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: true, removeAttachments: false));

            result.Undo.Should().NotBeNull();
            result.Undo.Moves.Should().BeEmpty();
            result.Undo.CreatedCopies.Should().ContainSingle().Which.Should().BeSameAs(copyRef);
            store.DidNotReceive().GetParentFolder(Arg.Any<MailItemRef>()); // not needed without a move
        }

        [Fact]
        public void A_totally_failed_classify_has_no_undo_plan()
        {
            var store = Substitute.For<IMailStore>(); // move returns null -> error
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: false));

            result.Undo.Should().BeNull();
        }

        [Fact]
        public void Undo_plan_counts_attachment_strips_for_the_unrecoverable_warning()
        {
            var store = StoreWithWorkingMove();
            store.GetParentFolder(Arg.Any<MailItemRef>()).Returns(Dest("src", "Inbox"));
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(true);
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true));

            result.Undo.AttachmentStrips.Should().Be(1);
        }

        [Fact]
        public void A_move_that_resolves_nothing_counts_as_an_error_and_processes_nothing()
        {
            var store = Substitute.For<IMailStore>(); // MoveItemToFolder returns null
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: false));

            result.ItemsProcessed.Should().Be(0);
            result.OriginalsMoved.Should().Be(0);
            result.Errors.Should().Be(1);
        }

        [Fact]
        public void Remove_attachments_strips_the_filed_copy_not_the_kept_original()
        {
            var store = Substitute.For<IMailStore>();
            var item = Item("e1");
            var filedCopy = new MailItemRef("s1", "copy-in-dest", "e1");
            store.CopyItemToFolder(item, D1).Returns(filedCopy);
            store.RemoveAttachments(filedCopy).Returns(true);
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
        public void Remove_attachments_strips_the_moved_item_when_not_keeping_a_copy()
        {
            var store = StoreWithWorkingMove();
            var item = Item("e1");
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(true);
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(
                new[] { item }, new[] { D1 },
                keepCopy: false, removeAttachments: true));

            // The filed item IS the moved original - stripped at its new home,
            // by its post-move reference.
            store.Received(1).RemoveAttachments(
                Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"));
        }

        [Fact]
        public void An_encrypted_filed_item_keeps_its_attachments_and_is_reported()
        {
            var store = StoreWithWorkingMove();
            // The store refuses to strip (encrypted) - returns false.
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(false);
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1") }, new[] { D1 },
                keepCopy: false, removeAttachments: true));

            result.EncryptedStripSkips.Should().Be(1);
            result.ItemsProcessed.Should().Be(1); // still filed, just not stripped
            result.Errors.Should().Be(0);
        }

        [Fact]
        public void Remove_attachments_off_strips_nothing()
        {
            var store = StoreWithWorkingMove();
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
            store.GetConversationSiblings(e1).Returns(new ConversationSiblings(new[] { e2, sibling }, new string[0]));
            store.GetConversationSiblings(e2).Returns(new ConversationSiblings(new[] { sibling }, new string[0]));
            var sut = new ClassifierService(store);

            var preflight = sut.Preflight(new[] { e1, e2 }, widenConversation: true);

            preflight.Items.Select(i => i.EntryId).Should().BeEquivalentTo(new[] { "e1", "e2", "sib" });
            preflight.Items.Should().HaveCount(3); // no duplicates despite the cross-reported sibling
        }

        [Fact]
        public void Preflight_reports_skipped_encrypted_siblings_without_filing_them()
        {
            var store = Substitute.For<IMailStore>();
            var e1 = Item("e1");
            var plainSibling = new MailItemRef("s1", "sib", "reply");
            store.GetConversationSiblings(e1).Returns(
                new ConversationSiblings(new[] { plainSibling }, new[] { "Secret thread" }));
            var sut = new ClassifierService(store);

            var preflight = sut.Preflight(new[] { e1 }, widenConversation: true);

            // The encrypted sibling is not in the set to file...
            preflight.Items.Select(i => i.EntryId).Should().BeEquivalentTo(new[] { "e1", "sib" });
            // ...but it is reported so the caller can warn.
            preflight.SkippedEncrypted.Should().ContainSingle().Which.Should().Be("Secret thread");
        }

        [Fact]
        public void Preflight_dedupes_skipped_encrypted_reports_across_source_items()
        {
            var store = Substitute.For<IMailStore>();
            var e1 = Item("e1");
            var e2 = Item("e2");
            // Both source items report the same encrypted sibling subject.
            store.GetConversationSiblings(e1).Returns(new ConversationSiblings(new MailItemRef[0], new[] { "Secret" }));
            store.GetConversationSiblings(e2).Returns(new ConversationSiblings(new MailItemRef[0], new[] { "Secret" }));
            var sut = new ClassifierService(store);

            var preflight = sut.Preflight(new[] { e1, e2 }, widenConversation: true);

            preflight.SkippedEncrypted.Should().ContainSingle().Which.Should().Be("Secret");
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
        public void Classify_marks_the_filed_item_complete_for_flagged_items_only()
        {
            var store = StoreWithWorkingMove();
            var flagged = Item("e1");
            var plain = Item("e2");
            store.IsFlaggedIncomplete(flagged).Returns(true);
            store.IsFlaggedIncomplete(plain).Returns(false);
            var sut = new ClassifierService(store);

            sut.Classify(new ClassifyRequest(
                new[] { flagged, plain }, new[] { D1 },
                keepCopy: false, removeAttachments: false, markTasksComplete: true));

            // Acts on the filed (moved) reference, never the pre-move original.
            store.Received(1).MarkTaskComplete(Arg.Is<MailItemRef>(m => m.EntryId == "moved-e1"));
            store.DidNotReceive().MarkTaskComplete(Arg.Is<MailItemRef>(m => m.EntryId == "moved-e2"));
            store.DidNotReceive().MarkTaskComplete(flagged);
        }

        [Fact]
        public void Classify_marks_nothing_complete_when_not_requested()
        {
            var store = StoreWithWorkingMove();
            var flagged = Item("e1");
            store.IsFlaggedIncomplete(flagged).Returns(true);
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
            store.DidNotReceive().MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>());
            store.DidNotReceive().DeleteItem(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void A_failing_destination_does_not_prevent_filing_into_the_others()
        {
            var store = StoreWithWorkingMove();
            var item = Item("e1");
            // D1 (e.g. a store root that refuses items) fails the copy; the
            // original must still move on to D2.
            store.When(s => s.CopyItemToFolder(item, D1))
                 .Do(_ => throw new InvalidOperationException("root refuses items"));
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { item }, new[] { D1, D2 },
                keepCopy: false, removeAttachments: false));

            store.Received(1).MoveItemToFolder(item, D2); // the good destination still got it
            result.CopiesMade.Should().Be(0);
            result.OriginalsMoved.Should().Be(1);
            result.Errors.Should().Be(1);                 // the failing D1
            result.ItemsProcessed.Should().Be(1);
        }

        [Fact]
        public void Original_stays_in_place_when_every_destination_fails()
        {
            var store = Substitute.For<IMailStore>();
            var item = Item("e1");
            store.When(s => s.CopyItemToFolder(item, Arg.Any<FolderNode>()))
                 .Do(_ => throw new InvalidOperationException("boom"));
            store.When(s => s.MoveItemToFolder(item, Arg.Any<FolderNode>()))
                 .Do(_ => throw new InvalidOperationException("boom"));
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { item }, new[] { D1, D2 },
                keepCopy: false, removeAttachments: false));

            store.DidNotReceive().DeleteItem(item); // no data loss on a total failure
            result.Errors.Should().Be(2);
            result.ItemsProcessed.Should().Be(0);
            result.OriginalsMoved.Should().Be(0);
        }

        [Fact]
        public void One_failing_item_is_counted_and_the_rest_still_process()
        {
            var store = StoreWithWorkingMove();
            // Fail the move for e1 only.
            store.When(s => s.MoveItemToFolder(Arg.Is<MailItemRef>(m => m.EntryId == "e1"), Arg.Any<FolderNode>()))
                 .Do(_ => throw new InvalidOperationException("boom"));
            var sut = new ClassifierService(store);

            var result = sut.Classify(new ClassifyRequest(
                new[] { Item("e1"), Item("e2") }, new[] { D1 },
                keepCopy: false, removeAttachments: false));

            result.Errors.Should().Be(1);
            result.ItemsProcessed.Should().Be(1); // e2
            result.OriginalsMoved.Should().Be(1); // only e2 moved
        }
    }
}
