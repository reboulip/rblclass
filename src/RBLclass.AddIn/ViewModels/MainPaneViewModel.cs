using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Threading;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// The single always-open RBLclass pane: one folder-search list that does
    /// both "open a folder" and "classify". A row's checkbox toggles it as a
    /// classify destination (single click); double-click - or Enter on the
    /// highlighted row - files the current Outlook mail selection into that one
    /// folder; the per-row open button navigates Outlook to it; the per-row "+"
    /// creates a sub-folder under it. "Classify to checked folders" files into
    /// every checked folder at once. Replaces the former Open/Classify split
    /// (built fresh rather than retrofitted - [[prefer-isolated-new-ui-over-retrofit]]).
    /// </summary>
    public sealed class MainPaneViewModel : ObservableObject
    {
        private readonly IFolderSearch _search;
        private readonly IClassifier _classifier;
        private readonly Func<IReadOnlyList<MailItemRef>> _getSelection;
        private readonly Func<FolderNode, string, FolderNode> _createSubfolder;
        private readonly Action<FolderNode, bool> _navigate;
        private readonly ISettingsStore _settings;
        private readonly Func<int, bool?> _confirmMarkTasksComplete;
        private readonly Func<string, string> _promptForName;

        private string _query = string.Empty;
        private bool _allResults, _keepCopy, _removeAttachments, _widenConversation, _openInNewWindow;
        private SelectableFolder _selectedResult;
        private string _selectionSummary = "No mail selected.";
        private string _status = string.Empty;
        private bool _isBusy;

        public MainPaneViewModel(
            IFolderSearch search,
            IClassifier classifier,
            Func<IReadOnlyList<MailItemRef>> getSelection,
            Func<FolderNode, string, FolderNode> createSubfolder = null,
            Action<FolderNode, bool> navigate = null,
            ISettingsStore settings = null,
            Func<int, bool?> confirmMarkTasksComplete = null,
            Func<string, string> promptForName = null)
        {
            _search = search ?? throw new ArgumentNullException(nameof(search));
            _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
            _getSelection = getSelection;
            _createSubfolder = createSubfolder;
            _navigate = navigate;
            _settings = settings;
            _confirmMarkTasksComplete = confirmMarkTasksComplete;
            _promptForName = promptForName;

            if (_settings != null)
            {
                _allResults = _settings.GetBool(SettingsKeys.AllResults, false);
                _keepCopy = _settings.GetBool(SettingsKeys.KeepCopy, false);
                _removeAttachments = _settings.GetBool(SettingsKeys.RemoveAttachments, false);
                _widenConversation = _settings.GetBool(SettingsKeys.WidenConversation, false);
                _openInNewWindow = _settings.GetBool(SettingsKeys.OpenInNewWindow, false);
            }
        }

        public ObservableCollection<SelectableFolder> Results { get; } =
            new ObservableCollection<SelectableFolder>();

        public string Query
        {
            get => _query;
            set { if (SetProperty(ref _query, value)) Refresh(); }
        }

        public bool AllResults
        {
            get => _allResults;
            set { if (SetProperty(ref _allResults, value)) { _settings?.SetBool(SettingsKeys.AllResults, value); Refresh(); } }
        }

        public bool KeepCopy
        {
            get => _keepCopy;
            set { if (SetProperty(ref _keepCopy, value)) _settings?.SetBool(SettingsKeys.KeepCopy, value); }
        }

        public bool RemoveAttachments
        {
            get => _removeAttachments;
            set { if (SetProperty(ref _removeAttachments, value)) _settings?.SetBool(SettingsKeys.RemoveAttachments, value); }
        }

        public bool WidenConversation
        {
            get => _widenConversation;
            set { if (SetProperty(ref _widenConversation, value)) _settings?.SetBool(SettingsKeys.WidenConversation, value); }
        }

        public bool OpenInNewWindow
        {
            get => _openInNewWindow;
            set { if (SetProperty(ref _openInNewWindow, value)) _settings?.SetBool(SettingsKeys.OpenInNewWindow, value); }
        }

        public SelectableFolder SelectedResult
        {
            get => _selectedResult;
            set => SetProperty(ref _selectedResult, value);
        }

        public string SelectionSummary
        {
            get => _selectionSummary;
            private set => SetProperty(ref _selectionSummary, value);
        }

        public string Status
        {
            get => _status;
            private set => SetProperty(ref _status, value);
        }

        /// <summary>True while a classify runs - disables the pane and guards against double-firing.</summary>
        public bool IsBusy
        {
            get => _isBusy;
            private set { if (SetProperty(ref _isBusy, value)) OnPropertyChanged(nameof(IsNotBusy)); }
        }

        public bool IsNotBusy => !_isBusy;

        public void SetSelectionCount(int count) =>
            SelectionSummary = count == 1 ? "1 mail selected." : count + " mails selected.";

        public void RefreshSelection() =>
            SetSelectionCount(_getSelection != null ? _getSelection().Count : 0);

        /// <summary>Navigate Outlook to a folder (the per-row open button).</summary>
        public void OpenFolder(FolderNode folder)
        {
            if (folder != null) _navigate?.Invoke(folder, _openInNewWindow);
        }

        /// <summary>File into the highlighted (or first) folder - Enter in the search box.</summary>
        public void FileToHighlighted()
        {
            var folder = (SelectedResult ?? (Results.Count > 0 ? Results[0] : null))?.Folder;
            if (folder == null) { Status = "No matching folder to file into."; return; }
            DoClassify(new[] { folder });
        }

        /// <summary>File into one specific folder - the double-clicked row.</summary>
        public void FileToFolder(FolderNode folder)
        {
            if (folder != null) DoClassify(new[] { folder });
        }

        /// <summary>File into every checked folder (the Classify button).</summary>
        public void ClassifyChecked()
        {
            var destinations = Results.Where(r => r.IsSelected).Select(r => r.Folder).ToList();
            if (destinations.Count == 0) { Status = "Check at least one destination folder."; return; }
            DoClassify(destinations);
        }

        /// <summary>Create a sub-folder under a specific folder (the per-row "+").</summary>
        public void CreateSubfolderUnder(FolderNode parent)
        {
            if (_isBusy || parent == null || _createSubfolder == null) return;

            string name = _promptForName != null ? _promptForName(parent.Name) : null;
            if (string.IsNullOrWhiteSpace(name)) return;

            var created = _createSubfolder(parent, name.Trim());
            if (created == null) { Status = "Could not create the folder."; return; }

            Status = "Created \"" + created.Name + "\" under " + parent.Name + ".";
            Refresh();
        }

        private void DoClassify(IReadOnlyList<FolderNode> destinations)
        {
            if (_isBusy) return; // ignore a repeat trigger while a classify is in flight

            var items = _getSelection != null ? _getSelection() : new MailItemRef[0];
            if (items.Count == 0) { Status = "Select one or more mails in Outlook first."; return; }

            IsBusy = true;
            Status = "Filing…";
            Dispatcher.CurrentDispatcher.Invoke(new Action(() => { }), DispatcherPriority.Render);

            try
            {
                var preflight = _classifier.Preflight(items, _widenConversation);

                bool markTasksComplete = false;
                if (preflight.FlaggedIncomplete.Count > 0 && _confirmMarkTasksComplete != null)
                {
                    var answer = _confirmMarkTasksComplete(preflight.FlaggedIncomplete.Count);
                    if (answer == null) { Status = "Classify cancelled."; return; }
                    markTasksComplete = answer.Value;
                }

                var result = _classifier.Classify(
                    new ClassifyRequest(preflight.Items, destinations, _keepCopy, _removeAttachments, markTasksComplete));

                string verb = _keepCopy ? "Copied" : "Filed";
                Status = verb + " " + result.ItemsProcessed + " mail(s) to " +
                         destinations.Count + " folder(s)" +
                         (result.Errors > 0 ? " (" + result.Errors + " failed)" : "") + ".";

                if (preflight.SkippedEncrypted.Count > 0)
                    Status += " " + preflight.SkippedEncrypted.Count +
                              " encrypted message(s) in the conversation were left in place.";

                RefreshSelection();
            }
            finally
            {
                Dispatcher.CurrentDispatcher.BeginInvoke(
                    new Action(() => IsBusy = false), DispatcherPriority.Background);
            }
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

            var outcome = _search.Search(_query, new FolderSearchOptions(matchMode, _allResults, maxResults));

            Results.Clear();
            foreach (var r in outcome.Results)
                Results.Add(new SelectableFolder(r));

            if (outcome.TotalMatchCount == 0)
                Status = string.IsNullOrWhiteSpace(_query) ? string.Empty : "No matching folders.";
            else if (outcome.LimitExceeded)
                Status = "Showing " + Results.Count + " of " + outcome.TotalMatchCount + " - refine your search.";
            else
                Status = outcome.TotalMatchCount + " folder(s).";
        }
    }
}
