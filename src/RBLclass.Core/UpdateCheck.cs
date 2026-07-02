using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Pure decision logic for the opt-in update notification (v2.7.0.0 E1):
    /// given the running version, a fetched "latest" string, and the last
    /// version the user dismissed, decide whether to surface the update banner.
    /// Kept in Core (no network, no COM) so it is fully unit-testable.
    /// </summary>
    public static class UpdateCheck
    {
        /// <summary>How long a dismissal snoozes the banner for the same version (v2.7.0.0 E1).</summary>
        public static readonly TimeSpan DismissWindow = TimeSpan.FromDays(7);

        /// <summary>
        /// True when <paramref name="remoteRaw"/> parses to a version newer than
        /// <paramref name="current"/> that is not currently snoozed. A dismissal of
        /// the same version (<paramref name="lastSeenRaw"/>, dismissed at
        /// <paramref name="dismissedUtc"/>) suppresses the banner only until
        /// <paramref name="window"/> has elapsed since; after that it re-appears. A
        /// remote version newer than the dismissed one always shows. On true,
        /// <paramref name="remote"/> is the version to show; on false it is null.
        /// </summary>
        public static bool ShouldNotify(Version current, string remoteRaw, string lastSeenRaw,
                                        DateTime? dismissedUtc, DateTime nowUtc, TimeSpan window,
                                        out Version remote)
        {
            remote = null;
            if (current == null || string.IsNullOrWhiteSpace(remoteRaw)) return false;

            Version parsed;
            if (!Version.TryParse(remoteRaw.Trim(), out parsed)) return false;
            if (parsed <= current) return false;

            // A dismissal only snoozes the exact version dismissed, and only for the
            // window. Once it elapses (or there is no timestamp), show it again.
            Version lastSeen;
            if (Version.TryParse((lastSeenRaw ?? string.Empty).Trim(), out lastSeen) && lastSeen >= parsed
                && dismissedUtc.HasValue && (nowUtc - dismissedUtc.Value) < window)
                return false;

            remote = parsed;
            return true;
        }
    }
}
