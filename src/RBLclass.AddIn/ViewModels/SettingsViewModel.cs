using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
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
        private string _maxResultsText;
        private string _minSearchLengthText;
        private string _searchDebounceMsText;

        public SettingsViewModel(ISettingsStore store)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _settings = Settings.Load(_store);
            _maxResultsText = _settings.MaxResults.ToString(CultureInfo.InvariantCulture);
            _minSearchLengthText = _settings.MinSearchLength.ToString(CultureInfo.InvariantCulture);
            _searchDebounceMsText = _settings.SearchDebounceMs.ToString(CultureInfo.InvariantCulture);
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
        public IReadOnlyList<TriageModeOption> TriageModes { get; } = new[]
        {
            new TriageModeOption(SentItemTriageMode.AskEveryTime, "Ask me each time"),
            new TriageModeOption(SentItemTriageMode.MoveToInbox, "Move to Inbox"),
            new TriageModeOption(SentItemTriageMode.Delete, "Delete"),
            new TriageModeOption(SentItemTriageMode.Leave, "Leave in Sent Items"),
        };

        public SentItemTriageMode SentItemTriageMode
        {
            get => _settings.SentItemTriageMode;
            set => Apply(_settings.SentItemTriageMode, value, v => _settings.SentItemTriageMode = v);
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
