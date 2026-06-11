using System;
using System.Collections.ObjectModel;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// View model for the "Open a folder" task pane: type-ahead keyword search
    /// over <see cref="IFolderSearch"/>, and navigate Outlook to a chosen folder.
    /// </summary>
    public sealed class FolderSearchViewModel : ObservableObject
    {
        private readonly IFolderSearch _search;
        private readonly Action<FolderNode, bool> _navigate;
        private readonly ISettingsStore _settings;

        private string _query = string.Empty;
        private bool _allResults;
        private bool _openInNewWindow; // default same-window (legacy settingsNewWindow default)
        private FolderSearchResult _selectedResult;
        private string _status = string.Empty;

        public FolderSearchViewModel(IFolderSearch search,
                                     Action<FolderNode, bool> navigate,
                                     ISettingsStore settings = null)
        {
            _search = search ?? throw new ArgumentNullException(nameof(search));
            _navigate = navigate;
            _settings = settings;

            // Load persisted toggles (default to the legacy defaults).
            if (_settings != null)
            {
                _allResults = _settings.GetBool(SettingsKeys.AllResults, false);
                _openInNewWindow = _settings.GetBool(SettingsKeys.OpenInNewWindow, false);
            }
        }

        public ObservableCollection<FolderSearchResult> Results { get; } =
            new ObservableCollection<FolderSearchResult>();

        public string Query
        {
            get => _query;
            set { if (SetProperty(ref _query, value)) Refresh(); }
        }

        public bool AllResults
        {
            get => _allResults;
            set
            {
                if (SetProperty(ref _allResults, value))
                {
                    _settings?.SetBool(SettingsKeys.AllResults, value);
                    Refresh();
                }
            }
        }

        public FolderSearchResult SelectedResult
        {
            get => _selectedResult;
            set => SetProperty(ref _selectedResult, value);
        }

        /// <summary>
        /// Whether navigation opens a new explorer window vs. retargeting the
        /// active one. In-memory for now; persisted to the Settings table in
        /// Step 9 (legacy settingsNewWindow).
        /// </summary>
        public bool OpenInNewWindow
        {
            get => _openInNewWindow;
            set
            {
                if (SetProperty(ref _openInNewWindow, value))
                    _settings?.SetBool(SettingsKeys.OpenInNewWindow, value);
            }
        }

        public string Status
        {
            get => _status;
            private set => SetProperty(ref _status, value);
        }

        /// <summary>Navigate to the highlighted result (double-click / Open button).</summary>
        public void OpenSelected() => Open(SelectedResult ?? FirstOrNull());

        /// <summary>Navigate to the first result (Enter in the search box).</summary>
        public void OpenFirst() => Open(FirstOrNull());

        private FolderSearchResult FirstOrNull() => Results.Count > 0 ? Results[0] : null;

        private void Open(FolderSearchResult result)
        {
            if (result != null)
                _navigate?.Invoke(result.Folder, _openInNewWindow);
        }

        private void Refresh()
        {
            var matchMode = FolderMatchMode.Substring;
            var maxResults = FolderSearchOptions.DefaultMaxResults;
            if (_settings != null)
            {
                var settings = Settings.Load(_settings);
                matchMode = settings.FolderMatchMode;
                maxResults = settings.MaxResults;
            }

            var options = new FolderSearchOptions(matchMode, _allResults, maxResults);
            var outcome = _search.Search(_query, options);

            Results.Clear();
            foreach (var r in outcome.Results)
                Results.Add(r);

            if (outcome.TotalMatchCount == 0)
                Status = string.IsNullOrWhiteSpace(_query) ? string.Empty : "No matching folders.";
            else if (outcome.LimitExceeded)
                Status = "Showing " + Results.Count + " of " + outcome.TotalMatchCount +
                         " - refine your search.";
            else
                Status = outcome.TotalMatchCount + " folder(s).";
        }
    }
}
