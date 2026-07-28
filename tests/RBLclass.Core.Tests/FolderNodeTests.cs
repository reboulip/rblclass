using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class FolderNodeTests
    {
        [Fact]
        public void StoreRoot_has_IsStoreRoot_true()
        {
            var folder = new FolderNode("s1", "root-entry", parentEntryId: null,
                "ArchiveStore", "ArchiveStore", isLeaf: false);

            folder.IsStoreRoot.Should().BeTrue();
        }

        [Fact]
        public void NonRoot_folder_has_IsStoreRoot_false()
        {
            var folder = new FolderNode("s1", "child-entry", parentEntryId: "root-entry",
                "Inbox", "ArchiveStore / Inbox", isLeaf: true);

            folder.IsStoreRoot.Should().BeFalse();
        }
    }
}
