using System;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class UpdateCheckTests
    {
        private static readonly Version Current = new Version(2, 7, 0, 0);
        private static readonly DateTime Now = new DateTime(2026, 7, 2, 12, 0, 0, DateTimeKind.Utc);
        private static readonly TimeSpan Week = TimeSpan.FromDays(7);

        private static bool Notify(string remoteRaw, string lastSeen, DateTime? dismissedUtc, out Version remote) =>
            UpdateCheck.ShouldNotify(Current, remoteRaw, lastSeen, dismissedUtc, Now, Week, out remote);

        [Fact]
        public void Newer_version_notifies_and_reports_it()
        {
            Notify("9.9.9.9", "", null, out var remote).Should().BeTrue();
            remote.Should().Be(new Version(9, 9, 9, 9));
        }

        [Fact]
        public void Same_version_does_not_notify()
        {
            Notify("2.7.0.0", "", null, out var remote).Should().BeFalse();
            remote.Should().BeNull();
        }

        [Fact]
        public void Older_version_does_not_notify()
        {
            Notify("2.0.0.0", "", null, out var remote).Should().BeFalse();
            remote.Should().BeNull();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not-a-version")]
        public void Missing_or_unparseable_remote_does_not_notify(string raw)
        {
            Notify(raw, "", null, out var remote).Should().BeFalse();
            remote.Should().BeNull();
        }

        [Fact]
        public void Recently_dismissed_version_is_snoozed()
        {
            Notify("9.9.9.9", "9.9.9.9", Now.AddDays(-1), out var remote).Should().BeFalse();
            remote.Should().BeNull();
        }

        [Fact]
        public void Dismissal_older_than_the_window_re_appears()
        {
            Notify("9.9.9.9", "9.9.9.9", Now.AddDays(-8), out var remote).Should().BeTrue();
            remote.Should().Be(new Version(9, 9, 9, 9));
        }

        [Fact]
        public void A_newer_version_than_the_dismissed_one_shows_within_the_window()
        {
            Notify("9.9.9.9", "9.0.0.0", Now.AddDays(-1), out var remote).Should().BeTrue();
            remote.Should().Be(new Version(9, 9, 9, 9));
        }
    }
}
