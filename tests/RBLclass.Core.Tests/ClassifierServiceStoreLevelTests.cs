using System.Linq;
using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// A protected mail filed at store level whose new EntryID is unknown comes
    /// back as an untracked ref (v2.14.0.0 A2, #62): it counts as filed, but
    /// nothing may resolve it afterwards.
    /// </summary>
    public class ClassifierServiceStoreLevelTests
    {
        private static readonly FolderNode Source = new FolderNode("s1", "src", null, "src", "Inbox", isLeaf: true);
        private static readonly FolderNode D1 = new FolderNode("s1", "d1", null, "d1", "Archive / A", isLeaf: true);
        private static readonly FolderNode D2 = new FolderNode("s1", "d2", null, "d2", "Archive / B", isLeaf: true);
        private static readonly FolderNode Trash = new FolderNode("s1", "trash", null, "trash", "Deleted", isLeaf: true);

        private static MailItemRef Item() => new MailItemRef("s1", "e1", "Secret subject");

        private static IMailStore ProtectedStore()
        {
            var store = Substitute.For<IMailStore>();
            store.IsEncryptedMail(Arg.Any<MailItemRef>()).Returns(true);
            store.GetParentFolder(Arg.Any<MailItemRef>()).Returns(Source);
            store.GetDeletedItemsFolder("s1").Returns(Trash);
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(_ => MailItemRef.Untracked("s1", "Secret subject"));
            store.CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(_ => MailItemRef.Untracked("s1", "Secret subject"));
            return store;
        }

        [Fact]
        public void Untracked_ref_has_no_entry_id()
        {
            var untracked = MailItemRef.Untracked("s1", "x");
            untracked.IsTracked.Should().BeFalse();
            untracked.EntryId.Should().BeEmpty();
            new MailItemRef("s1", "e1").IsTracked.Should().BeTrue();
        }

        [Fact]
        public void An_untracked_move_counts_as_filed_but_is_not_undoable()
        {
            var store = ProtectedStore();

            var result = new ClassifierService(store).Classify(
                new ClassifyRequest(new[] { Item() }, new[] { D1 }, new ClassifyOptions()));

            result.ItemsProcessed.Should().Be(1);
            result.OriginalsMoved.Should().Be(1);
            result.Errors.Should().Be(0);
            (result.Undo?.Moves ?? new UndoableMove[0]).Should().BeEmpty();
        }

        [Fact]
        public void An_untracked_move_is_never_resolved_again()
        {
            var store = ProtectedStore();

            new ClassifierService(store).Classify(
                new ClassifyRequest(new[] { Item() }, new[] { D1 }, new ClassifyOptions(safetyCopy: true)));

            // No safety copy from an unknown entry, no protection re-check of it.
            store.DidNotReceive().CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>());
            store.DidNotReceive().IsEncryptedMail(Arg.Is<MailItemRef>(r => !r.IsTracked));
            store.DidNotReceive().GetProtectionState(Arg.Is<MailItemRef>(r => !r.IsTracked));
        }

        [Fact]
        public void Untracked_copies_are_not_offered_for_undo()
        {
            var store = ProtectedStore();

            var result = new ClassifierService(store).Classify(
                new ClassifyRequest(new[] { Item() }, new[] { D1, D2 }, new ClassifyOptions(keepCopy: true)));

            result.CopiesMade.Should().Be(2);
            result.Errors.Should().Be(0);
            (result.Undo?.CreatedCopies ?? new MailItemRef[0]).Should().BeEmpty();
        }

        [Fact]
        public void A_tracked_store_level_move_stays_undoable()
        {
            var store = ProtectedStore();
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(_ => MailItemRef.VerifiedProtected("s1", "moved-e1", "Secret subject"));

            var result = new ClassifierService(store).Classify(
                new ClassifyRequest(new[] { Item() }, new[] { D1 }, new ClassifyOptions()));

            result.Undo.Moves.Select(m => m.Current.EntryId).Should().Equal("moved-e1");
        }

        [Fact]
        public void A_verified_filed_ref_is_not_rechecked_through_the_object_model()
        {
            var store = ProtectedStore();
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(_ => MailItemRef.VerifiedProtected("s1", "moved-e1", "Secret subject"));
            var log = Substitute.For<RBLclass.Core.Logging.IAppLog>();
            log.IsDebugEnabled.Returns(true);

            new ClassifierService(store, log: log).Classify(
                new ClassifyRequest(new[] { Item() }, new[] { D1 }, new ClassifyOptions()));

            store.DidNotReceive().IsEncryptedMail(Arg.Is<MailItemRef>(r => r.EntryId == "moved-e1"));
            store.DidNotReceive().GetProtectionState(Arg.Is<MailItemRef>(r => r.EntryId == "moved-e1"));
        }

        [Fact]
        public void A_plain_filed_ref_of_a_protected_mail_is_still_rechecked()
        {
            var store = ProtectedStore();
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(_ => new MailItemRef("s1", "moved-e1", "Secret subject"));

            new ClassifierService(store).Classify(
                new ClassifyRequest(new[] { Item() }, new[] { D1 }, new ClassifyOptions()));

            store.Received().IsEncryptedMail(Arg.Is<MailItemRef>(r => r.EntryId == "moved-e1"));
        }
    }
}
