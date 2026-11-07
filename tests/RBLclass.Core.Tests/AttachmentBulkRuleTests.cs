using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class AttachmentBulkRuleTests
    {
        private sealed class FakeRow : IAttachmentRowState
        {
            public AttachmentDispositionAction Action { get; set; } = AttachmentDispositionAction.SaveTo;
            public bool IsActionUserSet { get; set; }
            public bool IsInlineImage { get; set; }
            public string TargetDirectory { get; set; } = string.Empty;
        }

        private const AttachmentDispositionAction Delete = AttachmentDispositionAction.Delete;
        private const AttachmentDispositionAction SaveTo = AttachmentDispositionAction.SaveTo;
        private const AttachmentDispositionAction Keep = AttachmentDispositionAction.Keep;

        private static FakeRow UserSet(AttachmentDispositionAction action, string dir = "") =>
            new FakeRow { Action = action, IsActionUserSet = true, TargetDirectory = dir };

        private static FakeRow Untouched(AttachmentDispositionAction action = SaveTo, string dir = "") =>
            new FakeRow { Action = action, TargetDirectory = dir };

        /// <summary>What the grouped view does: apply the group choice to every unpinned row.</summary>
        private static void ApplyGroup(IEnumerable<FakeRow> rows, AttachmentDispositionAction action, string dir)
        {
            foreach (var r in rows.Where(AttachmentBulkRule.AppliesTo))
            {
                r.TargetDirectory = AttachmentBulkRule.DirectoryFor(r, action, dir);
                r.Action = action;
                r.IsActionUserSet = false;
            }
        }

        [Theory]
        [InlineData(Delete)]
        [InlineData(Keep)]
        public void A_user_set_delete_or_keep_row_is_pinned(AttachmentDispositionAction action)
        {
            AttachmentBulkRule.IsPinned(UserSet(action)).Should().BeTrue();
        }

        [Fact]
        public void A_user_set_save_to_row_is_not_pinned()
        {
            AttachmentBulkRule.IsPinned(UserSet(SaveTo)).Should().BeFalse();
        }

        [Theory]
        [InlineData(Delete)]
        [InlineData(SaveTo)]
        [InlineData(Keep)]
        public void An_untouched_row_is_not_pinned(AttachmentDispositionAction action)
        {
            AttachmentBulkRule.IsPinned(Untouched(action)).Should().BeFalse();
            AttachmentBulkRule.AppliesTo(Untouched(action)).Should().BeTrue();
        }

        [Theory]
        [InlineData(Delete)]
        [InlineData(SaveTo)]
        [InlineData(Keep)]
        public void An_inline_image_is_always_pinned(AttachmentDispositionAction action)
        {
            var row = new FakeRow { Action = action, IsInlineImage = true };

            AttachmentBulkRule.IsPinned(row).Should().BeTrue();
            AttachmentBulkRule.AppliesTo(row).Should().BeFalse();
        }

        [Fact]
        public void Save_to_takes_the_group_directory()
        {
            AttachmentBulkRule.DirectoryFor(Untouched(SaveTo, "A"), SaveTo, "D").Should().Be("D");
        }

        [Fact]
        public void Save_to_without_a_group_directory_keeps_the_row_directory()
        {
            AttachmentBulkRule.DirectoryFor(Untouched(SaveTo, "A"), SaveTo, "  ").Should().Be("A");
        }

        [Theory]
        [InlineData(Delete)]
        [InlineData(Keep)]
        public void Delete_and_keep_keep_the_row_directory(AttachmentDispositionAction action)
        {
            AttachmentBulkRule.DirectoryFor(Untouched(SaveTo, "A"), action, "D").Should().Be("A");
        }

        [Fact]
        public void The_shared_folder_goes_to_save_to_rows_only()
        {
            AttachmentBulkRule.SharedFolderFor(UserSet(SaveTo, "A"), "D").Should().Be("D");
            AttachmentBulkRule.SharedFolderFor(UserSet(Keep, "A"), "D").Should().Be("A");
            AttachmentBulkRule.SharedFolderFor(Untouched(Delete), "D").Should().BeEmpty();
        }

        [Fact]
        public void No_shared_folder_yet_keeps_the_row_folder()
        {
            AttachmentBulkRule.SharedFolderFor(UserSet(SaveTo, "A"), " ").Should().Be("A");
        }

        [Fact]
        public void Confirm_needs_a_directory_on_every_save_to_row()
        {
            AttachmentBulkRule.CanConfirm(new[] { Untouched(SaveTo, "A"), Untouched(SaveTo) }).Should().BeFalse();
        }

        [Fact]
        public void Delete_and_keep_rows_never_block_confirm()
        {
            AttachmentBulkRule.CanConfirm(new[] { UserSet(Keep), Untouched(Delete), Untouched(SaveTo, "A") })
                .Should().BeTrue();
        }

        [Fact]
        public void Confirm_is_allowed_with_no_rows()
        {
            AttachmentBulkRule.CanConfirm(new IAttachmentRowState[0]).Should().BeTrue();
        }

        [Fact]
        public void Seed_is_save_to_when_any_unpinned_row_saves()
        {
            AttachmentBulkRule.SeedGroupChoice(new[] { Untouched(Delete), Untouched(SaveTo, "A") }, out var dir)
                .Should().Be(SaveTo);
            dir.Should().Be("A");
        }

        [Fact]
        public void Seed_is_delete_only_when_every_unpinned_row_deletes()
        {
            AttachmentBulkRule.SeedGroupChoice(new[] { Untouched(Delete), Untouched(Delete) }, out _)
                .Should().Be(Delete);
        }

        [Fact]
        public void Seed_never_infers_delete_from_a_mix_of_delete_and_keep()
        {
            AttachmentBulkRule.SeedGroupChoice(new[] { Untouched(Delete), Untouched(Keep) }, out _)
                .Should().Be(Keep);
        }

        [Fact]
        public void Seed_ignores_pinned_rows()
        {
            var rows = new IAttachmentRowState[] { UserSet(Keep, "PINNED"), Untouched(Delete) };

            AttachmentBulkRule.SeedGroupChoice(rows, out var dir).Should().Be(Delete);
            dir.Should().BeEmpty();
        }

        [Fact]
        public void Seed_takes_the_first_directory_set()
        {
            AttachmentBulkRule.SeedGroupChoice(new[] { Untouched(SaveTo), Untouched(SaveTo, "B"), Untouched(SaveTo, "C") }, out var dir);

            dir.Should().Be("B");
        }

        [Fact]
        public void Seed_without_unpinned_rows_is_save_to_with_no_directory()
        {
            AttachmentBulkRule.SeedGroupChoice(new[] { UserSet(Keep, "A") }, out var dir).Should().Be(SaveTo);
            dir.Should().BeEmpty();
        }

        [Fact]
        public void A_group_save_to_never_flips_an_explicit_keep_or_delete()
        {
            var keep = UserSet(Keep, "A");
            var delete = UserSet(Delete);
            var save = UserSet(SaveTo, "A");
            var untouched = Untouched();

            ApplyGroup(new[] { keep, delete, save, untouched }, SaveTo, "D");

            keep.Action.Should().Be(Keep);
            keep.TargetDirectory.Should().Be("A");
            delete.Action.Should().Be(Delete);
            save.TargetDirectory.Should().Be("D");
            untouched.Action.Should().Be(SaveTo);
            untouched.TargetDirectory.Should().Be("D");
        }

        [Fact]
        public void An_exception_made_after_unticking_survives_ticking_again()
        {
            var rows = new[] { Untouched(), Untouched(), Untouched() };
            ApplyGroup(rows, Delete, "");

            // Unticked: rows keep the group's Delete, then the user keeps one.
            rows[1].Action = Keep;
            rows[1].IsActionUserSet = true;

            // Ticked again, the user picks Save to "D".
            ApplyGroup(rows, SaveTo, "D");

            rows[0].Action.Should().Be(SaveTo);
            rows[1].Action.Should().Be(Keep);
            rows[2].TargetDirectory.Should().Be("D");
        }
    }
}
