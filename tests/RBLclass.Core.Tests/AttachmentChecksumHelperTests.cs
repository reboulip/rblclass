using System;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class AttachmentChecksumHelperTests
    {
        private static Func<string, Stream> OpenFrom(Dictionary<string, byte[]> files) =>
            path => new MemoryStream(files[path]);

        [Fact]
        public void Identical_content_returns_true()
        {
            var files = new Dictionary<string, byte[]>
            {
                ["a"] = new byte[] { 1, 2, 3, 4, 5 },
                ["b"] = new byte[] { 1, 2, 3, 4, 5 },
            };
            AttachmentChecksumHelper.ContentEquals("a", "b", OpenFrom(files)).Should().BeTrue();
        }

        [Fact]
        public void Different_content_returns_false()
        {
            var files = new Dictionary<string, byte[]>
            {
                ["a"] = new byte[] { 1, 2, 3 },
                ["b"] = new byte[] { 1, 2, 4 },
            };
            AttachmentChecksumHelper.ContentEquals("a", "b", OpenFrom(files)).Should().BeFalse();
        }

        [Fact]
        public void Single_byte_difference_returns_false()
        {
            var files = new Dictionary<string, byte[]>
            {
                ["a"] = new byte[] { 0xAB },
                ["b"] = new byte[] { 0xAC },
            };
            AttachmentChecksumHelper.ContentEquals("a", "b", OpenFrom(files)).Should().BeFalse();
        }

        [Fact]
        public void Both_empty_files_are_equal()
        {
            var files = new Dictionary<string, byte[]>
            {
                ["a"] = new byte[0],
                ["b"] = new byte[0],
            };
            AttachmentChecksumHelper.ContentEquals("a", "b", OpenFrom(files)).Should().BeTrue();
        }

        [Fact]
        public void Empty_vs_nonempty_returns_false()
        {
            var files = new Dictionary<string, byte[]>
            {
                ["a"] = new byte[0],
                ["b"] = new byte[] { 1 },
            };
            AttachmentChecksumHelper.ContentEquals("a", "b", OpenFrom(files)).Should().BeFalse();
        }

        [Fact]
        public void Same_prefix_longer_file_returns_false()
        {
            var files = new Dictionary<string, byte[]>
            {
                ["a"] = new byte[] { 1, 2, 3 },
                ["b"] = new byte[] { 1, 2, 3, 4 },
            };
            AttachmentChecksumHelper.ContentEquals("a", "b", OpenFrom(files)).Should().BeFalse();
        }

        [Fact]
        public void Null_path1_throws_ArgumentNullException()
        {
            Action act = () => AttachmentChecksumHelper.ContentEquals(null, "b");
            act.Should().Throw<ArgumentNullException>().WithParameterName("path1");
        }

        [Fact]
        public void Null_path2_throws_ArgumentNullException()
        {
            Action act = () => AttachmentChecksumHelper.ContentEquals("a", null);
            act.Should().Throw<ArgumentNullException>().WithParameterName("path2");
        }
    }
}
