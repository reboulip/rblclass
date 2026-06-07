using System;
using System.Collections.ObjectModel;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// View model for the small modal folder-picker dialog used by the
    /// sent-item triage prompt's "Class…" action: type-ahead search over
    /// <see cref="IFolderSearch"/>, pick one folder and confirm. Deliberately
    /// separate from <see cref="ClassifyViewModel"/> - it drives a fixed item
    /// set rather than the live Outlook selection, and keeping it standalone
    /// avoids teaching the working Classify pane a second selection model.
    /// </summary>
    public sealed class FolderPickerViewModel : ObservableObject
    {
        private readonly IFolderSearch _search;

        private string _query = string.Empty;
        private FolderSearchResult _selectedResult;
        private string _status = "Type to search folders.";

        public FolderPickerViewModel(IFolderSearch search)
        {
            _search = search ?? throw new ArgumentNullException(nameof(search));
        }

        public ObservableCollection<FolderSearchResult> Results { get; } =
            new ObservableCollection<FolderSearchResult>();

        public string Query
        {
            get => _query;
            set { if (SetProperty(ref _query, value)) Refresh(); }
        }

        public FolderSearchResult SelectedResult
        {
            get => _selectedResult;
            set => SetProperty(ref _selectedResult, value);
        }

        public string Status
        {
            get => _status;
            private set => SetProperty(ref _status, value);
        }

        /// <summary>The folder to file into: the highlighted result, or the first match (Enter).</summary>
        public FolderNode ChosenFolder =>
            (SelectedResult ?? (Results.Count > 0 ? Results[0] : null))?.Folder;

        private void Refresh()
        {
            var outcome = _search.Search(_query, new FolderSearchOptions(FolderMatchMode.WordPrefix, allResults: false));

            Results.Clear();
            foreach (var r in outcome.Results)
                Results.Add(r);

            Status = Results.Count == 0
                ? (string.IsNullOrWhiteSpace(_query) ? "Type to search folders." : "No matching folders.")
                : outcome.TotalMatchCount + " folder(s).";
        }
    }
}
