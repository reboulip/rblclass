using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Typed, validated snapshot of every user-facing option (legacy §7),
    /// loaded from and saved to an <see cref="ISettingsStore"/> as one unit.
    /// This is what the Step 9 settings dialog binds to - it centralises the
    /// defaults/parsing/clamping that would otherwise be repeated at each
    /// scattered <c>GetBool</c>/<c>Get</c> call site.
    /// </summary>
    public sealed class Settings
    {
        /// <summary>Legacy hard-coded forgotten-attachment keyword list.</summary>
        public const string DefaultForgottenAttachmentKeywords = "attach;enclos;joint;PJ";

        /// <summary>
        /// Default pause after the last keystroke before folder search fires.
        /// ~100 ms is a typical inter-keystroke gap when typing a word, so
        /// 200 ms lets a burst finish yet still feels immediate after a stop.
        /// </summary>
        public const int DefaultSearchDebounceMs = 200;

        /// <summary>Upper clamp for <see cref="SearchDebounceMs"/> - beyond this the pane just feels broken.</summary>
        public const int MaxSearchDebounceMs = 2000;

        /// <summary>Upper clamp for <see cref="MinSearchLength"/>.</summary>
        public const int MaxMinSearchLength = 10;

        public bool OpenInNewWindow { get; set; }
        public bool AllResults { get; set; }
        public FolderMatchMode FolderMatchMode { get; set; }
        public int MaxResults { get; set; }
        public int MinSearchLength { get; set; }
        public int SearchDebounceMs { get; set; }
        public bool KeepCopy { get; set; }
        public bool RemoveAttachments { get; set; }
        public bool ClassifySafetyCopy { get; set; }
        public bool WidenConversation { get; set; }
        public bool SendExternalWarning { get; set; }
        public IReadOnlyList<string> InternalDomains { get; set; }
        public IReadOnlyList<string> ForgottenAttachmentKeywords { get; set; }
        public SentItemTriageMode SentItemTriageMode { get; set; }

        /// <summary>Read every key, falling back to the same defaults the individual call sites use today.</summary>
        public static Settings Load(ISettingsStore store)
        {
            return new Settings
            {
                OpenInNewWindow = store.GetBool(SettingsKeys.OpenInNewWindow, false),
                AllResults = store.GetBool(SettingsKeys.AllResults, false),
                FolderMatchMode = ParseMatchMode(store.Get(SettingsKeys.FolderMatchMode, null)),
                MaxResults = ParseMaxResults(store.Get(SettingsKeys.MaxResults, null)),
                MinSearchLength = ParseClampedInt(store.Get(SettingsKeys.MinSearchLength, null),
                    FolderSearchOptions.DefaultMinQueryLength, 1, MaxMinSearchLength),
                SearchDebounceMs = ParseClampedInt(store.Get(SettingsKeys.SearchDebounceMs, null),
                    DefaultSearchDebounceMs, 0, MaxSearchDebounceMs),
                KeepCopy = store.GetBool(SettingsKeys.KeepCopy, false),
                RemoveAttachments = store.GetBool(SettingsKeys.RemoveAttachments, false),
                ClassifySafetyCopy = store.GetBool(SettingsKeys.ClassifySafetyCopy, false),
                WidenConversation = store.GetBool(SettingsKeys.WidenConversation, false),
                SendExternalWarning = store.GetBool(SettingsKeys.SendExternalWarning, true),
                InternalDomains = ParseList(store.Get(SettingsKeys.InternalDomains, string.Empty)),
                ForgottenAttachmentKeywords = ParseList(store.Get(
                    SettingsKeys.ForgottenAttachmentKeywords, DefaultForgottenAttachmentKeywords)),
                SentItemTriageMode = ParseTriageMode(
                    store.Get(SettingsKeys.SentItemTriageMode, null),
                    store.GetBool(SettingsKeys.SentItemTriagePrompt, true))
            };
        }

        /// <summary>Write every key back as one unit (the dialog calls this on every change - eleven local key/value rows is not a cost worth optimising).</summary>
        public void Save(ISettingsStore store)
        {
            store.SetBool(SettingsKeys.OpenInNewWindow, OpenInNewWindow);
            store.SetBool(SettingsKeys.AllResults, AllResults);
            store.Set(SettingsKeys.FolderMatchMode, FolderMatchMode.ToString());
            store.Set(SettingsKeys.MaxResults, MaxResults.ToString(CultureInfo.InvariantCulture));
            store.Set(SettingsKeys.MinSearchLength, MinSearchLength.ToString(CultureInfo.InvariantCulture));
            store.Set(SettingsKeys.SearchDebounceMs, SearchDebounceMs.ToString(CultureInfo.InvariantCulture));
            store.SetBool(SettingsKeys.KeepCopy, KeepCopy);
            store.SetBool(SettingsKeys.RemoveAttachments, RemoveAttachments);
            store.SetBool(SettingsKeys.ClassifySafetyCopy, ClassifySafetyCopy);
            store.SetBool(SettingsKeys.WidenConversation, WidenConversation);
            store.SetBool(SettingsKeys.SendExternalWarning, SendExternalWarning);
            store.Set(SettingsKeys.InternalDomains, FormatList(InternalDomains));
            store.Set(SettingsKeys.ForgottenAttachmentKeywords, FormatList(ForgottenAttachmentKeywords));
            store.Set(SettingsKeys.SentItemTriageMode, SentItemTriageMode.ToString());
        }

        private static SentItemTriageMode ParseTriageMode(string raw, bool legacyPromptOn)
        {
            SentItemTriageMode mode;
            if (System.Enum.TryParse(raw, out mode) && System.Enum.IsDefined(typeof(SentItemTriageMode), mode))
                return mode;

            // No stored mode: migrate the legacy on/off prompt (on -> ask, off ->
            // leave). Fresh installs default to asking each time.
            return legacyPromptOn ? SentItemTriageMode.AskEveryTime : SentItemTriageMode.Leave;
        }

        private static FolderMatchMode ParseMatchMode(string raw)
        {
            FolderMatchMode mode;
            // Default (and fallback for an unrecognised value) is the broader
            // "contains" search; word-prefix is opt-in.
            return Enum.TryParse(raw, out mode) ? mode : FolderMatchMode.Substring;
        }

        private static int ParseMaxResults(string raw)
        {
            int value;
            if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) || value < 1)
                return FolderSearchOptions.DefaultMaxResults;
            return value;
        }

        /// <summary>Unparseable values fall back to the default; out-of-range values clamp to the nearest bound.</summary>
        private static int ParseClampedInt(string raw, int fallback, int min, int max)
        {
            int value;
            if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return fallback;
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        /// <summary>
        /// Split a semicolon-separated value into trimmed, non-empty entries
        /// (mirrors the legacy list format). Public so the settings dialog can
        /// translate its free-text editors to/from the same representation.
        /// </summary>
        public static IReadOnlyList<string> ParseList(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new string[0];
            return raw.Split(';')
                      .Select(s => s.Trim())
                      .Where(s => s.Length > 0)
                      .ToArray();
        }

        /// <summary>Inverse of <see cref="ParseList"/> - joins entries back into the stored/edited form.</summary>
        public static string FormatList(IEnumerable<string> values)
        {
            if (values == null) return string.Empty;
            return string.Join(";", values.Select(s => s.Trim()).Where(s => s.Length > 0));
        }
    }
}
