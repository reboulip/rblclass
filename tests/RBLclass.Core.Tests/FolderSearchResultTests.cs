using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// The hierarchical-display split (v2.4 C1): FullPath is broken into its
    /// segments on <see cref="FolderNode.PathSeparator"/> (" / "), and the
    /// collapsed suffix never leaks into the segments.
    /// </summary>
    public sealed class FolderSearchResultTests
    {
        private static FolderNode Node(string fullPath) =>
            new FolderNode("s1", "e1", null, "leaf", fullPath, isLeaf: true);

        [Fact]
        public void PathSegments_root_level_folder_is_single_segment()
        {
            var r = new FolderSearchResult(Node("Inbox"), isCollapsed: false);
            r.PathSegments.Should().Equal("Inbox");
        }

        [Fact]
        public void PathSegments_single_level()
        {
            var r = new FolderSearchResult(Node("Archive / Projects"), isCollapsed: false);
            r.PathSegments.Should().Equal("Archive", "Projects");
        }

        [Fact]
        public void PathSegments_multi_level()
        {
            var r = new FolderSearchResult(
                Node("Personal Archive / Projects / 2024 / RBLclass"), isCollapsed: false);
            r.PathSegments.Should().Equal("Personal Archive", "Projects", "2024", "RBLclass");
        }

        [Fact]
        public void PathSegments_path_ending_with_separator_keeps_trailing_empty_segment()
        {
            var r = new FolderSearchResult(Node("Archive / "), isCollapsed: false);
            r.PathSegments.Should().Equal("Archive", "");
        }

        [Fact]
        public void PathSegments_excludes_collapsed_suffix_even_when_collapsed()
        {
            var r = new FolderSearchResult(Node("Archive / Projects"), isCollapsed: true);
            r.DisplayPath.Should().EndWith(FolderSearchResult.CollapsedSuffix);
            r.PathSegments.Should().Equal("Archive", "Projects");
        }
    }
}
