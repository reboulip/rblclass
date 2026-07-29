using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Records when the most recent SelectionChange event arrived and decides
    /// whether the burst has "settled" (no new event for at least
    /// <see cref="DebounceMs"/> milliseconds). The caller owns the actual timer
    /// scheduling (DispatcherTimer in the add-in shell); this class only tracks
    /// signal times so the settle logic is unit-testable via an injectable clock.
    /// (v2.10.0.0 C1)
    /// </summary>
    public sealed class SelectionChangeDebouncer
    {
        private readonly Func<DateTime> _clock;
        private DateTime? _lastSignalAt;

        /// <summary>
        /// How many milliseconds of silence after the last <see cref="Signal"/>
        /// before <see cref="HasSettled"/> returns true. Clamped to zero if negative.
        /// </summary>
        public int DebounceMs { get; }

        /// <param name="debounceMs">Settle window in milliseconds (typically <see cref="Settings.DefaultSearchDebounceMs"/>).</param>
        /// <param name="clock">UTC clock; defaults to <see cref="DateTime.UtcNow"/> when null (for production use).</param>
        public SelectionChangeDebouncer(int debounceMs, Func<DateTime> clock = null)
        {
            DebounceMs = debounceMs < 0 ? 0 : debounceMs;
            _clock = clock ?? (() => DateTime.UtcNow);
        }

        /// <summary>Record that a new SelectionChange event just fired.</summary>
        public void Signal()
        {
            _lastSignalAt = _clock();
        }

        /// <summary>
        /// True when at least <see cref="DebounceMs"/> milliseconds have elapsed
        /// since the last <see cref="Signal"/> call. False if no Signal was ever
        /// received or the settle window has not elapsed yet.
        /// </summary>
        public bool HasSettled()
        {
            if (_lastSignalAt == null) return false;
            return (_clock() - _lastSignalAt.Value).TotalMilliseconds >= DebounceMs;
        }

        /// <summary>Clear settle state; call after the debounced action fires.</summary>
        public void Reset()
        {
            _lastSignalAt = null;
        }
    }
}
