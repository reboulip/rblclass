using System;
using System.Diagnostics;
using System.Threading;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class AttachmentDefaultFolderResolverTests
    {
        private static readonly TimeSpan Timeout = TimeSpan.FromMilliseconds(500);

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void A_blank_setting_resolves_to_empty_without_probing(string configured)
        {
            bool probed = false;

            AttachmentDefaultFolderResolver.Resolve(configured, _ => probed = true, Timeout).Should().BeEmpty();
            probed.Should().BeFalse();
        }

        [Fact]
        public void A_reachable_folder_is_returned_trimmed()
        {
            AttachmentDefaultFolderResolver.Resolve(@"  C:\Pièces jointes  ", _ => true, Timeout)
                .Should().Be(@"C:\Pièces jointes");
        }

        [Fact]
        public void A_missing_folder_resolves_to_empty()
        {
            AttachmentDefaultFolderResolver.Resolve(@"C:\Gone", _ => false, Timeout).Should().BeEmpty();
        }

        [Fact]
        public void A_failing_probe_resolves_to_empty()
        {
            AttachmentDefaultFolderResolver.Resolve(@"\\server\share", _ => throw new InvalidOperationException(), Timeout)
                .Should().BeEmpty();
        }

        [Fact]
        public void A_probe_that_hangs_resolves_to_empty_after_the_timeout()
        {
            var watch = Stopwatch.StartNew();

            var result = AttachmentDefaultFolderResolver.Resolve(@"\\dead\share",
                _ => { Thread.Sleep(5000); return true; }, TimeSpan.FromMilliseconds(100));

            result.Should().BeEmpty();
            watch.ElapsedMilliseconds.Should().BeLessThan(2000);
        }
    }
}
