using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class SaveAsPlannerTests
    {
        [Theory]
        [InlineData(SaveAsFormat.Msg, null, SaveAsFormat.Msg)]
        [InlineData(SaveAsFormat.Msg, SaveAsFormat.Pdf, SaveAsFormat.Msg)] // a stale picker answer never overrides a fixed setting
        [InlineData(SaveAsFormat.Pdf, null, SaveAsFormat.Pdf)]
        [InlineData(SaveAsFormat.Pdf, SaveAsFormat.Msg, SaveAsFormat.Pdf)]
        [InlineData(SaveAsFormat.Ask, SaveAsFormat.Pdf, SaveAsFormat.Pdf)]
        [InlineData(SaveAsFormat.Ask, SaveAsFormat.Msg, SaveAsFormat.Msg)]
        [InlineData(SaveAsFormat.Ask, null, SaveAsFormat.Msg)] // no answer: the historical format
        public void Resolve_picks_the_concrete_format(SaveAsFormat setting, SaveAsFormat? choice, SaveAsFormat expected)
        {
            SaveAsPlanner.Resolve(setting, choice).Should().Be(expected);
        }

        [Fact]
        public void PdfPath_uses_the_sanitized_subject_with_a_pdf_extension()
        {
            SaveAsPlanner.PdfPath(@"C:\out", "Quarterly report", p => false)
                .Should().Be(@"C:\out\Quarterly report.pdf");
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void PdfPath_falls_back_to_email_for_a_blank_subject(string name)
        {
            SaveAsPlanner.PdfPath(@"C:\out", name, p => false).Should().Be(@"C:\out\email.pdf");
        }

        [Fact]
        public void PdfPath_suffixes_on_collision()
        {
            var taken = new HashSet<string> { @"C:\out\Report.pdf", @"C:\out\Report (1).pdf" };

            SaveAsPlanner.PdfPath(@"C:\out", "Report", taken.Contains)
                .Should().Be(@"C:\out\Report (2).pdf");
        }

        [Fact]
        public void Partition_splits_protected_from_plain_and_keeps_order()
        {
            var items = new[] { "a", "S1", "b", "S2", "c" };

            var split = SaveAsPlanner.Partition(items, s => s.StartsWith("S"));

            split.Plain.Should().Equal("a", "b", "c");
            split.Protected.Should().Equal("S1", "S2");
        }

        [Fact]
        public void Partition_of_an_all_plain_batch_has_no_protected_mail()
        {
            SaveAsPlanner.Partition(new[] { 1, 2, 3 }, i => false).Protected.Should().BeEmpty();
        }

        [Fact]
        public void Partition_of_an_empty_batch_is_empty()
        {
            var split = SaveAsPlanner.Partition(Enumerable.Empty<int>(), i => true);

            split.Plain.Should().BeEmpty();
            split.Protected.Should().BeEmpty();
        }
    }
}
