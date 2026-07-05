using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using RBLclass.Core.Logging;

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

        /// <summary>Default retention window for the auto-classify conversation history (days).</summary>
        public const int DefaultAutoClassHistoryDays = 90;

        /// <summary>Upper clamp — 3650 days (≈10 years) is functionally "unlimited".</summary>
        public const int MaxAutoClassHistoryDays = 3650;

        public bool OpenInNewWindow { get; set; }
        public bool AllResults { get; set; }
        public FolderMatchMode FolderMatchMode { get; set; }
        public int MaxResults { get; set; }
        public int MinSearchLength { get; set; }
        public int SearchDebounceMs { get; set; }
        public bool KeepCopy { get; set; }
        public bool RemoveAttachments { get; set; }
        public bool ClassifySafetyCopy { get; set; }
        public string ExternalBannerSignature { get; set; }
        public bool StripBannerOnReply { get; set; }
        public bool StripBannerOnClassify { get; set; }
        public bool StripBannerOnAutoClassify { get; set; }
        public bool WidenConversation { get; set; }
        public bool SendExternalWarning { get; set; }
        public IReadOnlyList<string> InternalDomains { get; set; }
        public IReadOnlyList<string> ForgottenAttachmentKeywords { get; set; }
        public SentItemTriageMode SentItemTriageMode { get; set; }
        public bool ClassifyAfterMoveToInbox { get; set; }
        public string SentItemTriageArchiveStoreId { get; set; }
        public string SentItemTriageArchiveEntryId { get; set; }
        public string SentItemTriageArchivePath { get; set; }
        public IReadOnlyList<string> AttachmentFavoriteFolders { get; set; }
        public AttachmentRemovalMode AttachmentRemovalMode { get; set; }
        public AttachmentLabelLocation AttachmentLabelLocation { get; set; }
        public SearchPathDisplayMode SearchPathDisplayMode { get; set; }
        public int AutoClassHistoryDays { get; set; }
        public bool ClassifyMeetingItems { get; set; }
        public bool MarkAsReadOnClassify { get; set; }
        public bool UpdateCheckEnabled { get; set; }
        public string UpdateCheckUrl { get; set; }
        public string UpdateCheckLastSeenVersion { get; set; }
        public string UpdateCheckDismissedUtc { get; set; }
        public bool ShowPaneOnStartup { get; set; }
        public bool StripSavedAttachments { get; set; }
        public string PreferredUiLanguage { get; set; }
        public bool DiagnosticLogging { get; set; }

        /// <summary>
        /// Read every key in one snapshot (a single connection via
        /// <see cref="ISettingsStore.GetAll"/>, v2.7.1.0 E2), falling back to the
        /// same defaults the individual call sites use today.
        /// </summary>
        public static Settings Load(ISettingsStore store)
        {
            var snap = store.GetAll();

            string Str(string key, string def = null)
                => snap.TryGetValue(key, out var v) && v != null ? v : def;

            bool Bool(string key, bool def)
            {
                var raw = Str(key);
                if (raw == null) return def;
                if (raw == "1") return true;
                if (raw == "0") return false;
                return bool.TryParse(raw, out var p) ? p : def;
            }

            return new Settings
            {
                OpenInNewWindow = Bool(SettingsKeys.OpenInNewWindow, false),
                AllResults = Bool(SettingsKeys.AllResults, false),
                FolderMatchMode = ParseMatchMode(Str(SettingsKeys.FolderMatchMode)),
                MaxResults = ParseMaxResults(Str(SettingsKeys.MaxResults)),
                MinSearchLength = ParseClampedInt(Str(SettingsKeys.MinSearchLength),
                    FolderSearchOptions.DefaultMinQueryLength, 1, MaxMinSearchLength),
                SearchDebounceMs = ParseClampedInt(Str(SettingsKeys.SearchDebounceMs),
                    DefaultSearchDebounceMs, 0, MaxSearchDebounceMs),
                KeepCopy = Bool(SettingsKeys.KeepCopy, false),
                RemoveAttachments = Bool(SettingsKeys.RemoveAttachments, false),
                ClassifySafetyCopy = Bool(SettingsKeys.ClassifySafetyCopy, false),
                ExternalBannerSignature = Str(SettingsKeys.ExternalBannerSignature, string.Empty),
                StripBannerOnReply = Bool(SettingsKeys.StripBannerOnReply, false),
                StripBannerOnClassify = Bool(SettingsKeys.StripBannerOnClassify, false),
                StripBannerOnAutoClassify = Bool(SettingsKeys.StripBannerOnAutoClassify, false),
                WidenConversation = Bool(SettingsKeys.WidenConversation, false),
                SendExternalWarning = Bool(SettingsKeys.SendExternalWarning, true),
                InternalDomains = ParseList(Str(SettingsKeys.InternalDomains, string.Empty)),
                ForgottenAttachmentKeywords = ParseList(Str(
                    SettingsKeys.ForgottenAttachmentKeywords, DefaultForgottenAttachmentKeywords)),
                SentItemTriageMode = ParseTriageMode(
                    Str(SettingsKeys.SentItemTriageMode),
                    Bool(SettingsKeys.SentItemTriagePrompt, true)),
                ClassifyAfterMoveToInbox = Bool(SettingsKeys.ClassifyAfterMoveToInbox, true),
                SentItemTriageArchiveStoreId = Str(SettingsKeys.SentItemTriageArchiveStoreId, string.Empty),
                SentItemTriageArchiveEntryId = Str(SettingsKeys.SentItemTriageArchiveEntryId, string.Empty),
                SentItemTriageArchivePath = Str(SettingsKeys.SentItemTriageArchivePath, string.Empty),
                AttachmentFavoriteFolders = ParseList(Str(SettingsKeys.AttachmentFavoriteFolders, string.Empty)),
                AttachmentRemovalMode = ParseAttachmentRemovalMode(Str(SettingsKeys.AttachmentRemovalMode)),
                AttachmentLabelLocation = ParseAttachmentLabelLocation(Str(SettingsKeys.AttachmentLabelLocation)),
                SearchPathDisplayMode = ParsePathDisplayMode(
                    Str(SettingsKeys.SearchPathDisplayMode),
                    Bool(SettingsKeys.AutoExpandResults, false)),
                AutoClassHistoryDays = ParseClampedInt(Str(SettingsKeys.AutoClassHistoryDays),
                    DefaultAutoClassHistoryDays, 1, MaxAutoClassHistoryDays),
                ClassifyMeetingItems = Bool(SettingsKeys.ClassifyMeetingItems, false),
                MarkAsReadOnClassify = Bool(SettingsKeys.MarkAsReadOnClassify, false),
                UpdateCheckEnabled = Bool(SettingsKeys.UpdateCheckEnabled, false),
                UpdateCheckUrl = Str(SettingsKeys.UpdateCheckUrl, string.Empty),
                UpdateCheckLastSeenVersion = Str(SettingsKeys.UpdateCheckLastSeenVersion, string.Empty),
                UpdateCheckDismissedUtc = Str(SettingsKeys.UpdateCheckDismissedUtc, string.Empty),
                ShowPaneOnStartup = Bool(SettingsKeys.ShowPaneOnStartup, true),
                StripSavedAttachments = Bool(SettingsKeys.StripSavedAttachments, true),
                PreferredUiLanguage = ParseUiLanguage(Str(SettingsKeys.PreferredUiLanguage)),
                DiagnosticLogging = Bool(SettingsKeys.DiagnosticLogging, false)
            };
        }

        /// <summary>Write every key back as one batch transaction (a single connection, v2.7.1.0 E2).</summary>
        public void Save(ISettingsStore store)
        {
            store.SetMany(new[]
            {
                Pair(SettingsKeys.OpenInNewWindow,              OpenInNewWindow ? "1" : "0"),
                Pair(SettingsKeys.AllResults,                   AllResults ? "1" : "0"),
                Pair(SettingsKeys.FolderMatchMode,              FolderMatchMode.ToString()),
                Pair(SettingsKeys.MaxResults,                   MaxResults.ToString(CultureInfo.InvariantCulture)),
                Pair(SettingsKeys.MinSearchLength,              MinSearchLength.ToString(CultureInfo.InvariantCulture)),
                Pair(SettingsKeys.SearchDebounceMs,             SearchDebounceMs.ToString(CultureInfo.InvariantCulture)),
                Pair(SettingsKeys.KeepCopy,                     KeepCopy ? "1" : "0"),
                Pair(SettingsKeys.RemoveAttachments,            RemoveAttachments ? "1" : "0"),
                Pair(SettingsKeys.ClassifySafetyCopy,           ClassifySafetyCopy ? "1" : "0"),
                Pair(SettingsKeys.ExternalBannerSignature,      ExternalBannerSignature ?? string.Empty),
                Pair(SettingsKeys.StripBannerOnReply,           StripBannerOnReply ? "1" : "0"),
                Pair(SettingsKeys.StripBannerOnClassify,        StripBannerOnClassify ? "1" : "0"),
                Pair(SettingsKeys.StripBannerOnAutoClassify,    StripBannerOnAutoClassify ? "1" : "0"),
                Pair(SettingsKeys.WidenConversation,            WidenConversation ? "1" : "0"),
                Pair(SettingsKeys.SendExternalWarning,          SendExternalWarning ? "1" : "0"),
                Pair(SettingsKeys.InternalDomains,              FormatList(InternalDomains)),
                Pair(SettingsKeys.ForgottenAttachmentKeywords,  FormatList(ForgottenAttachmentKeywords)),
                Pair(SettingsKeys.SentItemTriageMode,           SentItemTriageMode.ToString()),
                Pair(SettingsKeys.ClassifyAfterMoveToInbox,     ClassifyAfterMoveToInbox ? "1" : "0"),
                Pair(SettingsKeys.SentItemTriageArchiveStoreId, SentItemTriageArchiveStoreId ?? string.Empty),
                Pair(SettingsKeys.SentItemTriageArchiveEntryId, SentItemTriageArchiveEntryId ?? string.Empty),
                Pair(SettingsKeys.SentItemTriageArchivePath,    SentItemTriageArchivePath ?? string.Empty),
                Pair(SettingsKeys.AttachmentFavoriteFolders,    FormatList(AttachmentFavoriteFolders)),
                Pair(SettingsKeys.AttachmentRemovalMode,        AttachmentRemovalMode.ToString()),
                Pair(SettingsKeys.AttachmentLabelLocation,      AttachmentLabelLocation.ToString()),
                Pair(SettingsKeys.SearchPathDisplayMode,        SearchPathDisplayMode.ToString()),
                Pair(SettingsKeys.AutoClassHistoryDays,         AutoClassHistoryDays.ToString(CultureInfo.InvariantCulture)),
                Pair(SettingsKeys.ClassifyMeetingItems,         ClassifyMeetingItems ? "1" : "0"),
                Pair(SettingsKeys.MarkAsReadOnClassify,         MarkAsReadOnClassify ? "1" : "0"),
                Pair(SettingsKeys.UpdateCheckEnabled,           UpdateCheckEnabled ? "1" : "0"),
                Pair(SettingsKeys.UpdateCheckUrl,               UpdateCheckUrl ?? string.Empty),
                Pair(SettingsKeys.UpdateCheckLastSeenVersion,   UpdateCheckLastSeenVersion ?? string.Empty),
                Pair(SettingsKeys.UpdateCheckDismissedUtc,      UpdateCheckDismissedUtc ?? string.Empty),
                Pair(SettingsKeys.ShowPaneOnStartup,            ShowPaneOnStartup ? "1" : "0"),
                Pair(SettingsKeys.StripSavedAttachments,        StripSavedAttachments ? "1" : "0"),
                Pair(SettingsKeys.PreferredUiLanguage,          PreferredUiLanguage),
                Pair(SettingsKeys.DiagnosticLogging,            DiagnosticLogging ? "1" : "0"),
            });
        }

        /// <summary>
        /// Same as <see cref="Save(ISettingsStore)"/>, plus one audit line
        /// summarising the security-relevant settings (v2.8.0.0 logging audit) -
        /// never the banner signature value. One line per Save, not per keystroke.
        /// </summary>
        public void Save(ISettingsStore store, IAppLog log)
        {
            Save(store);
            log?.Information(
                "Settings saved: sendExternalWarning={SendExternalWarning}, internalDomains={InternalDomains}, " +
                "classifySafetyCopy={ClassifySafetyCopy}, archiveTarget={ArchiveTarget}, forgottenAttachmentKeywords={ForgottenAttachmentKeywords}",
                SendExternalWarning, FormatList(InternalDomains), ClassifySafetyCopy,
                string.IsNullOrEmpty(SentItemTriageArchivePath) ? "(none)" : SentItemTriageArchivePath,
                FormatList(ForgottenAttachmentKeywords));
        }

        private static KeyValuePair<string, string> Pair(string key, string value)
            => new KeyValuePair<string, string>(key, value);

        private static SentItemTriageMode ParseTriageMode(string raw, bool legacyPromptOn)
        {
            SentItemTriageMode mode;
            if (System.Enum.TryParse(raw, out mode) && System.Enum.IsDefined(typeof(SentItemTriageMode), mode))
                return mode;

            // No stored mode: migrate the legacy on/off prompt (on -> ask, off ->
            // leave). Fresh installs default to asking each time.
            return legacyPromptOn ? SentItemTriageMode.AskEveryTime : SentItemTriageMode.Leave;
        }

        /// <summary>Unrecognised or missing values fall back to Modal (show the disposition prompt).</summary>
        private static AttachmentRemovalMode ParseAttachmentRemovalMode(string raw)
        {
            AttachmentRemovalMode mode;
            return Enum.TryParse(raw, out mode)
                   && Enum.IsDefined(typeof(AttachmentRemovalMode), mode)
                ? mode : AttachmentRemovalMode.Modal;
        }

        /// <summary>
        /// Parse the stored path-display mode; when absent, migrate the legacy
        /// AutoExpandResults flag (on → PerSegment, off → OneLine). Unrecognised
        /// values fall back to OneLine (v2.7.0.0 C1).
        /// </summary>
        private static SearchPathDisplayMode ParsePathDisplayMode(string raw, bool legacyAutoExpand)
        {
            SearchPathDisplayMode mode;
            if (Enum.TryParse(raw, out mode) && Enum.IsDefined(typeof(SearchPathDisplayMode), mode))
                return mode;
            return legacyAutoExpand ? SearchPathDisplayMode.PerSegment : SearchPathDisplayMode.OneLine;
        }

        /// <summary>Unrecognised or missing values fall back to Body (the reliable label location).</summary>
        private static AttachmentLabelLocation ParseAttachmentLabelLocation(string raw)
        {
            AttachmentLabelLocation loc;
            return Enum.TryParse(raw, out loc)
                   && Enum.IsDefined(typeof(AttachmentLabelLocation), loc)
                ? loc : AttachmentLabelLocation.Body;
        }

        /// <summary>Unrecognised or missing values fall back to "Auto" (follow Outlook's UI language).</summary>
        private static string ParseUiLanguage(string raw)
        {
            return UiLanguageResolver.IsSupportedLanguage(raw) || raw == "Auto" ? raw : "Auto";
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
