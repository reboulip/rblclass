using System;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class SelectionChangeDebouncerTests
    {
        [Fact]
        public void HasSettled_returns_false_before_any_signal()
        {
            DateTime t = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var sut = new SelectionChangeDebouncer(200, () => t);

            sut.HasSettled().Should().BeFalse();
        }

        [Fact]
        public void HasSettled_returns_false_immediately_after_signal()
        {
            DateTime t = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var sut = new SelectionChangeDebouncer(200, () => t);

            sut.Signal();

            sut.HasSettled().Should().BeFalse();
        }

        [Fact]
        public void HasSettled_returns_true_after_debounce_period_elapses()
        {
            DateTime t = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var sut = new SelectionChangeDebouncer(200, () => t);

            sut.Signal();
            t = t.AddMilliseconds(200);

            sut.HasSettled().Should().BeTrue();
        }

        [Fact]
        public void Second_signal_during_window_restarts_the_settle_clock()
        {
            DateTime t = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var sut = new SelectionChangeDebouncer(200, () => t);

            sut.Signal();
            t = t.AddMilliseconds(100);
            sut.Signal();
            t = t.AddMilliseconds(100);

            sut.HasSettled().Should().BeFalse("second signal reset the settle clock");
        }

        [Fact]
        public void Reset_clears_state_so_HasSettled_returns_false_after_reset()
        {
            DateTime t = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var sut = new SelectionChangeDebouncer(200, () => t);

            sut.Signal();
            t = t.AddMilliseconds(300);
            sut.HasSettled().Should().BeTrue();

            sut.Reset();

            sut.HasSettled().Should().BeFalse("Reset cleared the signal timestamp");
        }

        [Fact]
        public void Zero_debounce_settles_on_the_same_clock_tick_as_signal()
        {
            DateTime t = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var sut = new SelectionChangeDebouncer(0, () => t);

            sut.Signal();

            sut.HasSettled().Should().BeTrue();
        }

        [Fact]
        public void DebounceMs_property_reflects_constructor_value()
        {
            var sut = new SelectionChangeDebouncer(350);

            sut.DebounceMs.Should().Be(350);
        }

        [Fact]
        public void Negative_debounceMs_is_clamped_to_zero()
        {
            var sut = new SelectionChangeDebouncer(-1);

            sut.DebounceMs.Should().Be(0);
        }
    }
}
