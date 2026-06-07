using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// View model for the Classify pane: search destination folders (toggled
    /// with checkboxes), and file the current Outlook mail selection into the
    /// checked folders, with the legacy toggles (keep a copy, remove
    /// attachments, all results) persisted.
    /// </summary>
    public sealed class ClassifyViewModel : ObservableObject
    {
        private readonly IFolderSearch _search;
        private readonly IClassifier _classifier;
        private readonly Func<IReadOnlyList<MailItemRef>> _getSelection;
        private readonly Func<FolderNode, string, FolderNode> _createSubfolder;
        private readonly ISettingsStore _settings;
        private readonly Func<int, bool?> _confirmMarkTasksComplete;

        private string _query = string.Empty;
        private bool _allResults;
        private bool _keepCopy;
        private bool _removeAttachments;
        private bool _widenConversation;
        private string _selectionSummary = "No mail selected.";
        private string _status = string.Empty;

        public ClassifyViewModel(IFolderSearch search,
                                 IClassifier classifier,
                                 Func<IReadOnlyList<MailItemRef>> getSelection,
                                 Func<FolderNode, string, FolderNode> createSubfolder = null,
                                 ISettingsStore settings = null,
                                 Func<int, bool?> confirmMarkTasksComplete = null)
        {
            _search = search ?? throw new ArgumentNullException(nameof(search));
            _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
            _getSelection = getSelection;
            _createSubfolder = createSubfolder;
            _settings = settings;
            _confirmMarkTasksComplete = confirmMarkTasksComplete;

            if (_settings != null)
            {
                _allResults = _settings.GetBool(SettingsKeys.AllResults, false);
                _keepCopy = _settings.GetBool(SettingsKeys.KeepCopy, false);
                _removeAttachments = _settings.GetBool(SettingsKeys.RemoveAttachments, false);
                _widenConversation = _settings.GetBool(SettingsKeys.WidenConversation, false);
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
            set
            {
                if (SetProperty(ref _allResults, value))
                {
                    _settings?.SetBool(SettingsKeys.AllResults, value);
                    Refresh();
                }
            }
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

        /// <summary>Update the "N mails selected" label (from the live SelectionChange event).</summary>
        public void SetSelectionCount(int count)
        {
            SelectionSummary = count == 1 ? "1 mail selected." : count + " mails selected.";
        }

        /// <summary>Re-read the current Outlook mail selection (call when shown).</summary>
        public void RefreshSelection()
        {
            int count = _getSelection != null ? _getSelection().Count : 0;
            SetSelectionCount(count);
        }

        /// <summary>Classify into the checked destination folders (the Classify button).</summary>
        public void ClassifyChecked()
        {
            var destinations = Results.Where(r => r.IsSelected).Select(r => r.Folder).ToList();
            if (destinations.Count == 0)
            {
                Status = "Check at least one destination folder.";
                return;
            }
            DoClassify(destinations);
        }

        /// <summary>Classify into the first result folder (Enter in the search box).</summary>
        public void ClassifyToFirst()
        {
            if (Results.Count == 0)
            {
                Status = "No matching folder to file into.";
                return;
            }
            DoClassify(new[] { Results[0].Folder });
        }

        private void DoClassify(IReadOnlyList<FolderNode> destinations)
        {
            var items = _getSelection != null ? _getSelection() : new MailItemRef[0];
            if (items.Count == 0)
            {
                Status = "Select one or more mails in Outlook first.";
                return;
            }

            var preflight = _classifier.Preflight(items, _widenConversation);

            bool markTasksComplete = false;
            if (preflight.FlaggedIncomplete.Count > 0 && _confirmMarkTasksComplete != null)
            {
                var answer = _confirmMarkTasksComplete(preflight.FlaggedIncomplete.Count);
                if (answer == null)
                {
                    Status = "Classify cancelled.";
                    return;
                }
                markTasksComplete = answer.Value;
            }

            var result = _classifier.Classify(
                new ClassifyRequest(preflight.Items, destinations, _keepCopy, _removeAttachments, markTasksComplete));

            string verb = _keepCopy ? "Copied" : "Filed";
            Status = verb + " " + result.ItemsProcessed + " mail(s) to " +
                     destinations.Count + " folder(s)" +
                     (result.Errors > 0 ? " (" + result.Errors + " failed)" : "") + ".";

            RefreshSelection();
        }

        /// <summary>Create a sub-folder under the first checked folder, then re-run the search.</summary>
        public void CreateSubfolder(string name)
        {
            var parent = Results.FirstOrDefault(r => r.IsSelected)?.Folder;
            if (parent == null)
            {
                Status = "Check the parent folder first, then create the sub-folder.";
                return;
            }
            if (string.IsNullOrWhiteSpace(name))
            {
                Status = "Enter a name for the new folder.";
                return;
            }

            var created = _createSubfolder != null ? _createSubfolder(parent, name.Trim()) : null;
            if (created == null)
            {
                Status = "Could not create the folder.";
                return;
            }

            Status = "Created \"" + created.Name + "\" under " + parent.Name + ".";
            Refresh();
        }

        private void Refresh()
        {
            var options = new FolderSearchOptions(FolderMatchMode.WordPrefix, _allResults);
            var outcome = _search.Search(_query, options);

            Results.Clear();
            foreach (var r in outcome.Results)
                Results.Add(new SelectableFolder(r));
        }
    }
}
