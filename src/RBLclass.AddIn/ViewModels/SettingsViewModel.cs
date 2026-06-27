using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using RBLclass.AddIn.Localization;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// View model for the Step 9 settings dialog (legacy §7, modernised):
    /// every user-facing option bound directly to a <see cref="Settings"/>
    /// snapshot, persisted as a whole via <see cref="ISettingsStore"/> on
    /// every change - the same live-apply convention the Classify/folder-
    /// search panes already use, so "Close" is the only button and there is
    /// nothing to lose by walking away mid-edit.
    /// </summary>
    public sealed class SettingsViewModel : ObservableObject
    {
        private readonly ISettingsStore _store;
        private readonly Settings _settings;
        private readonly Func<string> _captureSelectedHtml;
        private readonly ILocalizationService _loc;
        private string _maxResultsText;
        private string _minSearchLengthText;
        private string _searchDebounceMsText;
        private string _autoClassHistoryDaysText;
        private string _bannerStatus;
        private string _bannerDiagnosticStatus = string.Empty;

        /// <param name="captureSelectedHtml">
        /// Returns the HTML body of the currently selected mail, for the "learn
        /// banner" capture (v2.2). Null/unset disables the Learn button.
        /// </param>
        public SettingsViewModel(ISettingsStore store, Func<string> captureSelectedHtml = null)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _captureSelectedHtml = captureSelectedHtml;
            _loc = TaskPaneServices.Localization;
            _settings = Settings.Load(_store);
            _maxResultsText = _settings.MaxResults.ToString(CultureInfo.InvariantCulture);
            _minSearchLengthText = _settings.MinSearchLength.ToString(CultureInfo.InvariantCulture);
            _searchDebounceMsText = _settings.SearchDebounceMs.ToString(CultureInfo.InvariantCulture);
            _autoClassHistoryDaysText = _settings.AutoClassHistoryDays.ToString(CultureInfo.InvariantCulture);
            _bannerStatus = DescribeBanner(_settings.ExternalBannerSignature);

            TriageModes = new[]
            {
                new TriageModeOption(SentItemTriageMode.AskEveryTime, _loc.GetString("Settings_TriageMode_AskEveryTime")),
                new TriageModeOption(SentItemTriageMode.MoveToInbox, _loc.GetString("Settings_TriageMode_MoveToInbox")),
                new TriageModeOption(SentItemTriageMode.Delete, _loc.GetString("Settings_TriageMode_Delete")),
                new TriageModeOption(SentItemTriageMode.Leave, _loc.GetString("Settings_TriageMode_Leave")),
            };

            UiLanguages = new[]
            {
                new UiLanguageOption("Auto", _loc.GetString("Settings_Language_Auto")),
                new UiLanguageOption("en", _loc.GetString("Settings_Language_English")),
                new UiLanguageOption("fr", _loc.GetString("Settings_Language_French")),
                new UiLanguageOption("de", _loc.GetString("Settings_Language_German")),
            };

            FavoriteFolders = new ObservableCollection<string>(
                _settings.AttachmentFavoriteFolders ?? new string[0]);
        }

        /// <summary>
        /// The user's favourite save-to directories (v2.4.0.0 F1), edited via the
        /// folder-browse dialog. Mutations persist immediately, like every other
        /// setting here.
        /// </summary>
        public ObservableCollection<string> FavoriteFolders { get; }

        /// <summary>F2: remove-attachments shows the disposition modal (Modal) or strips silently.</summary>
        public bool AttachmentRemovalModeIsModal
        {
            get => _settings.AttachmentRemovalMode == AttachmentRemovalMode.Modal;
            set { if (value) UpdateAttachmentRemovalMode(AttachmentRemovalMode.Modal); }
        }

        public bool AttachmentRemovalModeIsDeleteSilently
        {
            get => _settings.AttachmentRemovalMode == AttachmentRemovalMode.DeleteSilently;
            set { if (value) UpdateAttachmentRemovalMode(AttachmentRemovalMode.DeleteSilently); }
        }

        private void UpdateAttachmentRemovalMode(AttachmentRemovalMode mode)
        {
            if (_settings.AttachmentRemovalMode == mode) return;
            _settings.AttachmentRemovalMode = mode;
            _settings.Save(_store);
            OnPropertyChanged(nameof(AttachmentRemovalModeIsModal));
            OnPropertyChanged(nameof(AttachmentRemovalModeIsDeleteSilently));
        }

        /// <summary>F3: where the former-attachments label is recorded (Body, or InfoBar - deferred).</summary>
        public bool AttachmentLabelLocationIsBody
        {
            get => _settings.AttachmentLabelLocation == AttachmentLabelLocation.Body;
            set { if (value) UpdateAttachmentLabelLocation(AttachmentLabelLocation.Body); }
        }

        public bool AttachmentLabelLocationIsInfoBar
        {
            get => _settings.AttachmentLabelLocation == AttachmentLabelLocation.InfoBar;
            set { if (value) UpdateAttachmentLabelLocation(AttachmentLabelLocation.InfoBar); }
        }

        public bool AttachmentLabelLocationIsNone
        {
            get => _settings.AttachmentLabelLocation == AttachmentLabelLocation.None;
            set { if (value) UpdateAttachmentLabelLocation(AttachmentLabelLocation.None); }
        }

        private void UpdateAttachmentLabelLocation(AttachmentLabelLocation location)
        {
            if (_settings.AttachmentLabelLocation == location) return;
            _settings.AttachmentLabelLocation = location;
            _settings.Save(_store);
            OnPropertyChanged(nameof(AttachmentLabelLocationIsBody));
            OnPropertyChanged(nameof(AttachmentLabelLocationIsInfoBar));
            OnPropertyChanged(nameof(AttachmentLabelLocationIsNone));
        }

        /// <summary>Browse for a directory and add it (deduped, case-insensitive).</summary>
        public void AddFavoriteFolder()
        {
            string path = TaskPaneServices.BrowseForFolder?.Invoke();
            if (string.IsNullOrWhiteSpace(path)) return;
            if (FavoriteFolders.Contains(path, StringComparer.OrdinalIgnoreCase)) return;
            FavoriteFolders.Add(path);
            PersistFavorites();
        }

        /// <summary>Remove a favourite directory.</summary>
        public void RemoveFavoriteFolder(string path)
        {
            if (path == null || !FavoriteFolders.Remove(path)) return;
            PersistFavorites();
        }

        private void PersistFavorites()
        {
            _settings.AttachmentFavoriteFolders = FavoriteFolders.ToArray();
            _settings.Save(_store);
        }

        public bool OpenInNewWindow
        {
            get => _settings.OpenInNewWindow;
            set => Apply(_settings.OpenInNewWindow, value, v => _settings.OpenInNewWindow = v);
        }

        public bool AllResults
        {
            get => _settings.AllResults;
            set => Apply(_settings.AllResults, value, v => _settings.AllResults = v);
        }

        public bool AutoExpandResults
        {
            get => _settings.AutoExpandResults;
            set => Apply(_settings.AutoExpandResults, value, v => _settings.AutoExpandResults = v);
        }

        /// <summary>
        /// Word-prefix is the opt-in stricter mode; the default is the broader
        /// "contains" (substring) search. Bound to a single checkbox that is
        /// unchecked by default.
        /// </summary>
        public bool UseWordPrefixMatch
        {
            get => _settings.FolderMatchMode == FolderMatchMode.WordPrefix;
            set => Apply(UseWordPrefixMatch, value,
                         v => _settings.FolderMatchMode = v ? FolderMatchMode.WordPrefix : FolderMatchMode.Substring);
        }

        /// <summary>
        /// Free-text editor for <see cref="Settings.MaxResults"/>. Only commits
        /// (and persists) on a valid positive integer; an in-progress edit that
        /// doesn't yet parse is left alone rather than reverted or rejected.
        /// </summary>
        public string MaxResultsText
        {
            get => _maxResultsText;
            set
            {
                if (!SetProperty(ref _maxResultsText, value)) return;

                int parsed;
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) && parsed >= 1)
                {
                    _settings.MaxResults = parsed;
                    _settings.Save(_store);
                }
            }
        }

        /// <summary>
        /// Free-text editor for <see cref="Settings.MinSearchLength"/> (v2.2):
        /// search starts only once the query has this many characters. Commits
        /// on a valid value in range, like <see cref="MaxResultsText"/>.
        /// </summary>
        public string MinSearchLengthText
        {
            get => _minSearchLengthText;
            set
            {
                if (!SetProperty(ref _minSearchLengthText, value)) return;

                int parsed;
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed)
                    && parsed >= 1 && parsed <= Settings.MaxMinSearchLength)
                {
                    _settings.MinSearchLength = parsed;
                    _settings.Save(_store);
                }
            }
        }

        /// <summary>
        /// Free-text editor for <see cref="Settings.SearchDebounceMs"/> (v2.2):
        /// how long after the last keystroke the folder search fires.
        /// </summary>
        public string SearchDebounceMsText
        {
            get => _searchDebounceMsText;
            set
            {
                if (!SetProperty(ref _searchDebounceMsText, value)) return;

                int parsed;
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed)
                    && parsed >= 0 && parsed <= Settings.MaxSearchDebounceMs)
                {
                    _settings.SearchDebounceMs = parsed;
                    _settings.Save(_store);
                }
            }
        }

        public string AutoClassHistoryDaysText
        {
            get => _autoClassHistoryDaysText;
            set
            {
                if (!SetProperty(ref _autoClassHistoryDaysText, value)) return;

                int parsed;
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed)
                    && parsed >= 1 && parsed <= Settings.MaxAutoClassHistoryDays)
                {
                    _settings.AutoClassHistoryDays = parsed;
                    _settings.Save(_store);
                }
            }
        }

        public bool KeepCopy
        {
            get => _settings.KeepCopy;
            set => Apply(_settings.KeepCopy, value, v => _settings.KeepCopy = v);
        }

        public bool RemoveAttachments
        {
            get => _settings.RemoveAttachments;
            set => Apply(_settings.RemoveAttachments, value, v => _settings.RemoveAttachments = v);
        }

        /// <summary>
        /// Opt-in v2.2 guardrail: classify (without keep-a-copy) also leaves a
        /// copy of each filed mail in Deleted Items, like the old
        /// delete-after-copy behaviour did as a side effect.
        /// </summary>
        public bool ClassifySafetyCopy
        {
            get => _settings.ClassifySafetyCopy;
            set => Apply(_settings.ClassifySafetyCopy, value, v => _settings.ClassifySafetyCopy = v);
        }

        public bool WidenConversation
        {
            get => _settings.WidenConversation;
            set => Apply(_settings.WidenConversation, value, v => _settings.WidenConversation = v);
        }

        public bool SendExternalWarning
        {
            get => _settings.SendExternalWarning;
            set => Apply(_settings.SendExternalWarning, value, v => _settings.SendExternalWarning = v);
        }

        /// <summary>Auto-strip the learned banner from reply/forward drafts (default off).</summary>
        public bool StripBannerOnReply
        {
            get => _settings.StripBannerOnReply;
            set => Apply(_settings.StripBannerOnReply, value, v => _settings.StripBannerOnReply = v);
        }

        /// <summary>Default state of the classify-time "strip banner from the filed copy" tickbox.</summary>
        public bool StripBannerOnClassify
        {
            get => _settings.StripBannerOnClassify;
            set => Apply(_settings.StripBannerOnClassify, value, v => _settings.StripBannerOnClassify = v);
        }

        public bool StripBannerOnAutoClassify
        {
            get => _settings.StripBannerOnAutoClassify;
            set => Apply(_settings.StripBannerOnAutoClassify, value, v => _settings.StripBannerOnAutoClassify = v);
        }

        public string BannerDiagnosticStatus
        {
            get => _bannerDiagnosticStatus;
            private set => SetProperty(ref _bannerDiagnosticStatus, value);
        }

        /// <summary>True when a banner has been learned (gates the strip toggles in the view).</summary>
        public bool HasBanner => !string.IsNullOrWhiteSpace(_settings.ExternalBannerSignature);

        /// <summary>True when capture is wired (the Settings dialog was opened with a selection source).</summary>
        public bool CanLearnBanner => _captureSelectedHtml != null;

        /// <summary>Human-readable banner state, shown next to the Learn button.</summary>
        public string BannerStatus
        {
            get => _bannerStatus;
            private set => SetProperty(ref _bannerStatus, value);
        }

        /// <summary>
        /// Capture the external-sender banner from the currently selected mail:
        /// read its HTML body, extract the leading banner block, store it as the
        /// signature. Updates <see cref="BannerStatus"/> with the outcome.
        /// </summary>
        public void LearnBannerFromSelection()
        {
            if (_captureSelectedHtml == null) return;

            string html = _captureSelectedHtml();
            if (string.IsNullOrEmpty(html))
            {
                BannerStatus = _loc.GetString("Settings_Banner_SelectMailFirst");
                return;
            }

            string block = ExternalBannerStripper.ExtractBannerBlock(html);
            if (string.IsNullOrWhiteSpace(block))
            {
                BannerStatus = _loc.GetString("Settings_Banner_NotFound");
                return;
            }

            _settings.ExternalBannerSignature = block;
            _settings.Save(_store);
            BannerStatus = DescribeBanner(block);
            OnPropertyChanged(nameof(HasBanner));
        }

        public void DiagnoseSelectedMail()
        {
            if (_captureSelectedHtml == null)
            {
                BannerDiagnosticStatus = _loc.GetString("Settings_Banner_Diag_Unavailable");
                return;
            }

            string html = _captureSelectedHtml();
            string sig = _settings.ExternalBannerSignature;
            var result = ExternalBannerStripper.Diagnose(html, sig);

            switch (result.Outcome)
            {
                case BannerDiagnosticOutcome.NoSignature:
                    BannerDiagnosticStatus = _loc.GetString("Settings_Banner_Diag_NoSignature");
                    break;
                case BannerDiagnosticOutcome.NoSelection:
                    BannerDiagnosticStatus = _loc.GetString("Settings_Banner_Diag_NoSelection");
                    break;
                case BannerDiagnosticOutcome.Found:
                    BannerDiagnosticStatus = result.MatchedExact
                        ? _loc.GetString("Settings_Banner_Diag_FoundExact", result.SignatureLength)
                        : _loc.GetString("Settings_Banner_Diag_FoundTolerant", result.SignatureLength);
                    break;
                case BannerDiagnosticOutcome.NotFound:
                    BannerDiagnosticStatus = _loc.GetString("Settings_Banner_Diag_NotFound", result.SignatureLength);
                    break;
                default:
                    BannerDiagnosticStatus = string.Empty;
                    break;
            }
        }

        /// <summary>Forget the learned banner (and turn the strip options off, since they'd be inert).</summary>
        public void ClearBanner()
        {
            _settings.ExternalBannerSignature = string.Empty;
            _settings.StripBannerOnReply = false;
            _settings.StripBannerOnClassify = false;
            _settings.StripBannerOnAutoClassify = false;
            _settings.Save(_store);
            BannerStatus = DescribeBanner(string.Empty);
            BannerDiagnosticStatus = string.Empty;
            OnPropertyChanged(nameof(HasBanner));
            OnPropertyChanged(nameof(StripBannerOnReply));
            OnPropertyChanged(nameof(StripBannerOnClassify));
            OnPropertyChanged(nameof(StripBannerOnAutoClassify));
        }

        private string DescribeBanner(string signature) =>
            string.IsNullOrWhiteSpace(signature)
                ? _loc.GetString("Settings_Banner_NoneLearned")
                : _loc.GetString("Settings_Banner_Configured", signature.Trim().Length);

        /// <summary>Semicolon-separated domains treated as internal, edited as free text.</summary>
        public string InternalDomainsText
        {
            get => Settings.FormatList(_settings.InternalDomains);
            set => Apply(_settings.InternalDomains, Settings.ParseList(value),
                         v => _settings.InternalDomains = v);
        }

        /// <summary>Semicolon-separated forgotten-attachment keywords, edited as free text.</summary>
        public string ForgottenAttachmentKeywordsText
        {
            get => Settings.FormatList(_settings.ForgottenAttachmentKeywords);
            set => Apply(_settings.ForgottenAttachmentKeywords, Settings.ParseList(value),
                         v => _settings.ForgottenAttachmentKeywords = v);
        }

        /// <summary>Options for the sent-item triage dropdown (value + friendly label).</summary>
        public IReadOnlyList<TriageModeOption> TriageModes { get; }

        public SentItemTriageMode SentItemTriageMode
        {
            get => _settings.SentItemTriageMode;
            set => Apply(_settings.SentItemTriageMode, value, v => _settings.SentItemTriageMode = v);
        }

        /// <summary>
        /// After triage moves a sent mail to the Inbox, reveal the classify pane
        /// with that mail pinned so the user can file it immediately (E1).
        /// </summary>
        public bool ClassifyAfterMoveToInbox
        {
            get => _settings.ClassifyAfterMoveToInbox;
            set => Apply(_settings.ClassifyAfterMoveToInbox, value, v => _settings.ClassifyAfterMoveToInbox = v);
        }

        /// <summary>Options for the language dropdown (Auto/English/Français/Deutsch), Phase E.</summary>
        public IReadOnlyList<UiLanguageOption> UiLanguages { get; }

        /// <summary>
        /// "Auto" | "en" | "fr" | "de". Takes effect after restarting Outlook -
        /// the resolved language is read once at startup.
        /// </summary>
        public string PreferredUiLanguage
        {
            get => _settings.PreferredUiLanguage;
            set => Apply(_settings.PreferredUiLanguage, value, v => _settings.PreferredUiLanguage = v);
        }

        /// <summary>
        /// Change-detect, mutate the snapshot, notify and persist as a unit -
        /// the <see cref="Settings"/> equivalent of <see cref="ObservableObject.SetProperty"/>
        /// for state that lives in properties rather than ref-able fields.
        /// </summary>
        private void Apply<T>(T current, T value, Action<T> assign, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(current, value)) return;

            assign(value);
            OnPropertyChanged(propertyName);
            _settings.Save(_store);
        }
    }
}
