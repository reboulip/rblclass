using System;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class PathEllipsisTests
    {
        // Deterministic measurer: one unit per character (the ellipsis counts as
        // one character too), so widths are just string lengths.
        private static readonly Func<string, double> OnePerChar = s => s.Length;

        private const string E = PathEllipsis.DefaultEllipsis;

        [Fact]
        public void Returns_text_unchanged_when_it_already_fits()
        {
            PathEllipsis.TrimStart("Inbox", 10, OnePerChar).Should().Be("Inbox");
        }

        [Fact]
        public void Returns_text_unchanged_when_exactly_fits()
        {
            PathEllipsis.TrimStart("Inbox", 5, OnePerChar).Should().Be("Inbox");
        }

        [Fact]
        public void Keeps_the_longest_tail_with_a_leading_ellipsis()
        {
            // "A/B/C/Leaf" is 10 wide; at width 6 the longest fitting tail is
            // "/Leaf" (5) prefixed by the ellipsis (1) = 6.
            var result = PathEllipsis.TrimStart("A/B/C/Leaf", 6, OnePerChar);

            result.Should().Be(E + "/Leaf");
            result.Should().StartWith(E);
            OnePerChar(result).Should().BeLessThanOrEqualTo(6);
        }

        [Fact]
        public void One_more_tail_character_would_overflow()
        {
            // Confirms it is the *longest* fitting tail, not an over-trimmed one.
            var result = PathEllipsis.TrimStart("A/B/C/Leaf", 6, OnePerChar);
            (E + "C/Leaf").Length.Should().BeGreaterThan(6); // the next-longer tail
            result.Should().Be(E + "/Leaf");
        }

        [Fact]
        public void Returns_just_the_ellipsis_when_no_tail_character_fits()
        {
            PathEllipsis.TrimStart("abcdef", 1, OnePerChar).Should().Be(E);
        }

        [Fact]
        public void Empty_or_null_returns_empty()
        {
            PathEllipsis.TrimStart("", 10, OnePerChar).Should().Be("");
            PathEllipsis.TrimStart(null, 10, OnePerChar).Should().Be("");
        }

        [Fact]
        public void Non_positive_width_returns_text_unchanged()
        {
            PathEllipsis.TrimStart("abc", 0, OnePerChar).Should().Be("abc");
            PathEllipsis.TrimStart("abc", -5, OnePerChar).Should().Be("abc");
        }

        [Fact]
        public void Null_measure_throws()
        {
            Action act = () => PathEllipsis.TrimStart("abc", 10, null);
            act.Should().Throw<ArgumentNullException>();
        }
    }
}
