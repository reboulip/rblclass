using System.Collections.Generic;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class AttachmentFilenameCollisionResolverTests
    {
        [Fact]
        public void No_conflict_returns_the_original_path()
        {
            AttachmentFilenameCollisionResolver.Resolve(@"C:\out", "report.pdf", _ => false)
                .Should().Be(@"C:\out\report.pdf");
        }

        [Fact]
        public void First_collision_appends_1()
        {
            AttachmentFilenameCollisionResolver.Resolve(@"C:\out", "report.pdf",
                    p => p == @"C:\out\report.pdf")
                .Should().Be(@"C:\out\report (1).pdf");
        }

        [Fact]
        public void Increments_until_a_free_name()
        {
            var taken = new HashSet<string>
            {
                @"C:\out\report.pdf", @"C:\out\report (1).pdf", @"C:\out\report (2).pdf"
            };
            AttachmentFilenameCollisionResolver.Resolve(@"C:\out", "report.pdf", taken.Contains)
                .Should().Be(@"C:\out\report (3).pdf");
        }

        [Fact]
        public void Preserves_the_extension()
        {
            AttachmentFilenameCollisionResolver.Resolve(@"C:\out", "data.tar.gz",
                    p => p == @"C:\out\data.tar.gz")
                .Should().Be(@"C:\out\data.tar (1).gz");
        }

        [Fact]
        public void Handles_a_name_without_extension()
        {
            AttachmentFilenameCollisionResolver.Resolve(@"C:\out", "README",
                    p => p == @"C:\out\README")
                .Should().Be(@"C:\out\README (1)");
        }
    }
}
