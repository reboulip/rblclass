using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Threading;
using RBLclass.AddIn.Localization;
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
        private readonly Func<IReadOnlyList<FolderNode>> _getAllFolders;
        private readonly ILocalizationService _loc;

        private string _query = string.Empty;
        private bool _allResults, _keepCopy, _removeAttachments, _widenConversation, _stripBanner;
        private bool _canStripBanner;
        private bool _isOptionsExpanded;
        private SelectableFolder _selectedResult;
        private string _selectionSummary;
        private string _status = string.Empty;
        private bool _isBusy;
        private IndexStatus _indexStatus = IndexStatus.NotFound;

        // E1: a just-sent mail (moved to the Inbox by triage) pinned as the next
        // classify target, overriding the live Outlook selection until it is
        // filed or the user changes the real explorer selection.
        private MailItemRef _pinnedItem;

        // Debounces typing in the search box: re-search fires only once the
        // user has paused for Settings.SearchDebounceMs (v2.2).
        private DispatcherTimer _searchTimer;

        // Single-slot undo: the latest classify's reversal plan (v2.2). A new
        // classify overwrites it; executing it clears it.
        private ClassifyUndoPlan _lastUndo;

        public MainPaneViewModel(
            IFolderSearch search,
            IClassifier classifier,
            Func<IReadOnlyList<MailItemRef>> getSelection,
            Func<FolderNode, string, FolderNode> createSubfolder = null,
            Action<FolderNode, bool> navigate = null,
            ISettingsStore settings = null,
            Func<int, bool?> confirmMarkTasksComplete = null,
            Func<string, string> promptForName = null,
            Func<IReadOnlyList<FolderNode>> getAllFolders = null)
        {
            _search = search ?? throw new ArgumentNullException(nameof(search));
            _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
            _getSelection = getSelection;
            _createSubfolder = createSubfolder;
            _navigate = navigate;
            _settings = settings;
            _confirmMarkTasksComplete = confirmMarkTasksComplete;
            _promptForName = promptForName;
            _getAllFolders = getAllFolders;
            _loc = TaskPaneServices.Localization;
            _selectionSummary = _loc.GetString("Status_NoMailSelectedSummary");

            if (_settings != null)
            {
                _allResults = _settings.GetBool(SettingsKeys.AllResults, false);
                _keepCopy = _settings.GetBool(SettingsKeys.KeepCopy, false);
                _removeAttachments = _settings.GetBool(SettingsKeys.RemoveAttachments, false);
                _widenConversation = _settings.GetBool(SettingsKeys.WidenConversation, false);
                _stripBanner = _settings.GetBool(SettingsKeys.StripBannerOnClassify, false);
                _canStripBanner = !string.IsNullOrWhiteSpace(
                    _settings.Get(SettingsKeys.ExternalBannerSignature, string.Empty));
            }
        }

        public ObservableCollection<SelectableFolder> Results { get; } =
            new ObservableCollection<SelectableFolder>();

        /// <summary>
        /// Destinations accumulated across searches (v2.2): checking a row adds
        /// its folder here, and the set survives re-searches so multi-keyword
        /// destinations can be built up (search kw1, check, search kw2, check,
        /// classify once). Rendered as a removable chip strip; cleared after a
        /// successful "Classify to N folders".
        /// </summary>
        public ObservableCollection<FolderNode> SelectedDestinations { get; } =
            new ObservableCollection<FolderNode>();

        private readonly HashSet<string> _selectedKeys = new HashSet<string>();

        private static string KeyOf(FolderNode f) => f.StoreId + " " + f.EntryId;

        public bool HasSelectedDestinations => SelectedDestinations.Count > 0;

        public string SelectionHeader =>
            _loc.Plural(SelectedDestinations.Count, "Status_SelectedFolders_One", "Status_SelectedFolders_Other");

        public string ClassifyButtonText =>
            SelectedDestinations.Count == 0
                ? _loc.GetString("Status_ClassifyDefault")
                : _loc.Plural(SelectedDestinations.Count, "Status_ClassifyToFolders_One", "Status_ClassifyToFolders_Other");

        public string Query
        {
            get => _query;
            set { if (SetProperty(ref _query, value)) ScheduleRefresh(); }
        }

        // The five option toggles below are per-action overrides (v2.4 B1):
        // they are seeded from the persisted Settings defaults in the
        // constructor and reset back to them after each successful classify
        // (ResetOptionsToDefaults). Toggling one here does NOT write to
        // Settings - the Settings dialog is the only editor of the default.

        public bool AllResults
        {
            get => _allResults;
            set { if (SetProperty(ref _allResults, value)) Refresh(); }
        }

        public bool KeepCopy
        {
            get => _keepCopy;
            set => SetProperty(ref _keepCopy, value);
        }

        public bool RemoveAttachments
        {
            get => _removeAttachments;
            set => SetProperty(ref _removeAttachments, value);
        }

        public bool WidenConversation
        {
            get => _widenConversation;
            set => SetProperty(ref _widenConversation, value);
        }

        /// <summary>
        /// Strip the learned external banner from the filed copy (v2.2). The
        /// default comes from the StripBannerOnClassify setting; toggling it
        /// here is a per-action override (v2.4 B1) and does not change the
        /// setting. Only meaningful when a banner has been learned
        /// (<see cref="CanStripBanner"/>).
        /// </summary>
        public bool StripBanner
        {
            get => _stripBanner;
            set => SetProperty(ref _stripBanner, value);
        }

        /// <summary>True when a banner has been learned, so the strip tickbox is worth showing.</summary>
        public bool CanStripBanner
        {
            get => _canStripBanner;
            private set => SetProperty(ref _canStripBanner, value);
        }

        /// <summary>
        /// Whether the Options panel (the five per-action checkboxes) is expanded.
        /// Toggled by the Options button and by Tab from the query box; collapsed
        /// automatically after a successful classify (v2.4 B1).
        /// </summary>
        public bool IsOptionsExpanded
        {
            get => _isOptionsExpanded;
            set => SetProperty(ref _isOptionsExpanded, value);
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

        /// <summary>
        /// Folder-index build status, bound to the pane header's colored dot:
        /// Red = NotFound, Yellow = Indexing, Green = Ready. Driven by
        /// <see cref="IFolderIndexService"/> via <see cref="SubscribeToIndexStatus"/>.
        /// </summary>
        public IndexStatus IndexStatus
        {
            get => _indexStatus;
            private set { if (SetProperty(ref _indexStatus, value)) OnPropertyChanged(nameof(IndexStatusToolTip)); }
        }

        /// <summary>Localized tooltip for the index status dot.</summary>
        public string IndexStatusToolTip
        {
            get
            {
                switch (_indexStatus)
                {
                    case IndexStatus.Indexing: return _loc.GetString("IndexStatus_Indexing");
                    case IndexStatus.Ready: return _loc.GetString("IndexStatus_Ready");
                    default: return _loc.GetString("IndexStatus_NotFound");
                }
            }
        }

        /// <summary>
        /// Subscribe to the folder index service so <see cref="IndexStatus"/> and
        /// its tooltip track the walk lifecycle. Called once by the pane host.
        /// The service raises PropertyChanged on the Outlook STA thread, which is
        /// also the pane's WPF dispatcher thread, so binding updates are safe
        /// without marshalling.
        /// </summary>
        public void SubscribeToIndexStatus(IFolderIndexService service)
        {
            if (service == null) return;
            IndexStatus = service.IndexStatus;
            service.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(IFolderIndexService.IndexStatus))
                    IndexStatus = service.IndexStatus;
            };
        }

        /// <summary>True when the last classify can be undone (enables the Undo button).</summary>
        public bool CanUndo => _lastUndo != null;

        /// <summary>True when Auto-class is wired (the folder-index snapshot is available).</summary>
        public bool CanAutoClass => _getAllFolders != null;

        /// <summary>
        /// Auto-class the current selection (the pane's small Auto-class button):
        /// file each selected mail to its conversation's last recorded
        /// destination(s), show those folders in the results, and report what
        /// happened in the status line. No modal - everything lands in the pane.
        /// </summary>
        public void AutoClassSelected()
        {
            if (_isBusy || _getAllFolders == null) return;

            _searchTimer?.Stop(); // don't let a pending search overwrite our results

            var items = _getSelection != null ? _getSelection() : new MailItemRef[0];
            if (items.Count == 0) { Status = _loc.GetString("Status_NoMailSelectedAction"); return; }

            IsBusy = true;
            Status = _loc.GetString("Status_AutoClassing");
            Dispatcher.CurrentDispatcher.Invoke(new Action(() => { }), DispatcherPriority.Render);

            try
            {
                // One live-index snapshot, indexed for the staleness check.
                var byKey = new Dictionary<string, FolderNode>();
                foreach (var f in _getAllFolders())
                    byKey[KeyOf(f)] = f;
                Func<string, string, FolderNode> resolve = (storeId, entryId) =>
                {
                    FolderNode node;
                    return byKey.TryGetValue(storeId + " " + entryId, out node) ? node : null;
                };

                bool keepCopy = false, removeAttachments = false, safetyCopy = false;
                if (_settings != null)
                {
                    var s = Settings.Load(_settings);
                    keepCopy = s.KeepCopy;
                    removeAttachments = s.RemoveAttachments;
                    safetyCopy = s.ClassifySafetyCopy;
                }

                var result = _classifier.AutoClassify(items, resolve, keepCopy, removeAttachments, safetyCopy);

                // Show where the mail went (the destination folders), so the
                // user can see and open them; empty when nothing was filed.
                Results.Clear();
                foreach (var dest in result.FiledDestinations)
                    AddResultRow(new FolderSearchResult(dest, isCollapsed: false));

                Status = DescribeAutoClass(result, _loc);

                if (result.Undo != null)
                {
                    _lastUndo = result.Undo;
                    OnPropertyChanged(nameof(CanUndo));
                }

                RefreshSelection();
            }
            finally
            {
                Dispatcher.CurrentDispatcher.BeginInvoke(
                    new Action(() => IsBusy = false), DispatcherPriority.Background);
            }
        }

        private static string DescribeAutoClass(AutoClassifyResult r, ILocalizationService loc)
        {
            // Nothing filed and the only reason is "never classified before".
            if (r.Filed == 0 && r.NoHistory > 0 && r.StaleFolders == 0 && r.Errors == 0)
                return loc.GetString("Status_AutoClass_NoHistoryOnly");

            var parts = new List<string>();
            if (r.Filed > 0)
                parts.Add(loc.Plural(r.Filed, "Status_AutoClass_Filed_One", "Status_AutoClass_Filed_Other"));
            if (r.NoHistory > 0)
                parts.Add(loc.Plural(r.NoHistory, "Status_AutoClass_NoHistory_One", "Status_AutoClass_NoHistory_Other"));
            if (r.StaleFolders > 0)
                parts.Add(loc.Plural(r.StaleFolders, "Status_AutoClass_StaleFolders_One", "Status_AutoClass_StaleFolders_Other"));
            if (r.Errors > 0)
                parts.Add(loc.Plural(r.Errors, "Status_AutoClass_Errors_One", "Status_AutoClass_Errors_Other"));
            return parts.Count == 0 ? loc.GetString("Status_AutoClass_Nothing") : string.Join("; ", parts) + ".";
        }

        /// <summary>Reverse the last filing action completely (the Undo button).</summary>
        public void UndoLast()
        {
            if (_isBusy || _lastUndo == null) return;

            var plan = _lastUndo;
            _lastUndo = null;
            OnPropertyChanged(nameof(CanUndo));

            IsBusy = true;
            Status = _loc.GetString("Status_Undoing");
            Dispatcher.CurrentDispatcher.Invoke(new Action(() => { }), DispatcherPriority.Render);

            try
            {
                var undone = _classifier.Undo(plan);

                string movesRestored = _loc.Plural(undone.MovesRestored, "Status_Undo_MovesRestored_One", "Status_Undo_MovesRestored_Other");
                string copiesDeleted = _loc.Plural(undone.CopiesDeleted, "Status_Undo_CopiesDeleted_One", "Status_Undo_CopiesDeleted_Other");
                string stepsFailed = undone.Errors > 0
                    ? _loc.Plural(undone.Errors, "Status_Undo_StepsFailed_One", "Status_Undo_StepsFailed_Other")
                    : string.Empty;
                Status = _loc.GetString("Status_Undo_Result", movesRestored, copiesDeleted, stepsFailed);
                if (plan.AttachmentStrips > 0)
                    Status += _loc.GetString("Status_Undo_AttachmentsNotRestored");

                RefreshSelection();
            }
            finally
            {
                Dispatcher.CurrentDispatcher.BeginInvoke(
                    new Action(() => IsBusy = false), DispatcherPriority.Background);
            }
        }

        /// <summary>
        /// Push a live explorer selection count into the header. A genuine
        /// selection change drops a just-sent pin (E1) - the user has moved on.
        /// </summary>
        public void SetSelectionCount(int count)
        {
            ClearPin();
            SelectionSummary = _loc.Plural(count, "Status_MailSelected_One", "Status_MailSelected_Other");
        }

        /// <summary>
        /// Pin a just-sent mail as the next classify target, overriding the live
        /// selection until it is filed or the user changes the explorer selection
        /// (E1). A prominent banner shows it so the user knows they are filing the
        /// sent mail, not whatever is selected in the message list.
        /// </summary>
        public void PinMailForClassify(MailItemRef item)
        {
            if (item == null) return;
            _pinnedItem = item;
            SelectionSummary = string.Empty; // the pin banner takes over the header
            OnPropertyChanged(nameof(HasPinnedItem));
            OnPropertyChanged(nameof(PinnedItemSubject));
        }

        /// <summary>True while a just-sent mail is pinned (E1) - drives the pin banner.</summary>
        public bool HasPinnedItem => _pinnedItem != null;

        /// <summary>Subject of the pinned just-sent mail (E1), shown on its own wrapping line.</summary>
        public string PinnedItemSubject =>
            _pinnedItem == null
                ? string.Empty
                : (string.IsNullOrWhiteSpace(_pinnedItem.Subject)
                    ? _loc.GetString("SentTriage_NoSubject")
                    : _pinnedItem.Subject);

        private void ClearPin()
        {
            if (_pinnedItem == null) return;
            _pinnedItem = null;
            OnPropertyChanged(nameof(HasPinnedItem));
            OnPropertyChanged(nameof(PinnedItemSubject));
        }

        public void RefreshSelection()
        {
            // While a sent mail is pinned, keep the pin indicator - the show/theme
            // refresh must not overwrite it with the (unrelated) live count.
            if (_pinnedItem == null)
                SelectionSummary = _loc.Plural(_getSelection != null ? _getSelection().Count : 0,
                    "Status_MailSelected_One", "Status_MailSelected_Other");

            // A banner may have been learned (or forgotten) in Settings while the
            // pane was open - keep the strip tickbox's availability current.
            if (_settings != null)
                CanStripBanner = !string.IsNullOrWhiteSpace(
                    _settings.Get(SettingsKeys.ExternalBannerSignature, string.Empty));
        }

        /// <summary>
        /// Navigate Outlook to a folder (the per-row open button). New-window
        /// behaviour is read live from settings - since v2.2 it is configured
        /// only in the Settings dialog, no longer toggled in the pane.
        /// </summary>
        public void OpenFolder(FolderNode folder)
        {
            if (folder == null) return;
            bool newWindow = _settings != null && _settings.GetBool(SettingsKeys.OpenInNewWindow, false);
            _navigate?.Invoke(folder, newWindow);
        }

        /// <summary>File into the highlighted (or first) folder - Enter in the search box.</summary>
        public async Task FileToHighlighted()
        {
            FlushPendingSearch(); // never file based on stale, pre-debounce results
            var folder = (SelectedResult ?? (Results.Count > 0 ? Results[0] : null))?.Folder;
            if (folder == null) { Status = _loc.GetString("Status_NoMatchingFolderToFile"); return; }
            await DoClassifyAsync(new[] { folder });
        }

        /// <summary>File into one specific folder - the double-clicked row.</summary>
        public async Task FileToFolder(FolderNode folder)
        {
            if (folder != null) await DoClassifyAsync(new[] { folder });
        }

        /// <summary>File into every accumulated destination (the Classify button).</summary>
        public async Task ClassifyChecked()
        {
            var destinations = SelectedDestinations.ToList();
            if (destinations.Count == 0) { Status = _loc.GetString("Status_CheckAtLeastOneDestination"); return; }
            if (await DoClassifyAsync(destinations))
                ClearDestinations(); // the batch is filed - start the next one clean
        }

        /// <summary>Remove one accumulated destination (a chip's ✕, or unchecking its row).</summary>
        public void RemoveDestination(FolderNode folder)
        {
            if (folder == null || !_selectedKeys.Remove(KeyOf(folder))) return;

            for (int i = SelectedDestinations.Count - 1; i >= 0; i--)
            {
                if (KeyOf(SelectedDestinations[i]) == KeyOf(folder))
                    SelectedDestinations.RemoveAt(i);
            }

            // Uncheck the matching visible row, if any (its change handler
            // re-enters here and no-ops - the key is already gone).
            foreach (var row in Results)
            {
                if (KeyOf(row.Folder) == KeyOf(folder))
                    row.IsSelected = false;
            }

            NotifySelectionMeta();
        }

        /// <summary>Drop the whole accumulated selection (the clear-all link, or after classify).</summary>
        public void ClearDestinations()
        {
            _selectedKeys.Clear();
            SelectedDestinations.Clear();
            foreach (var row in Results)
                row.IsSelected = false; // re-entrant removes no-op on the empty set
            NotifySelectionMeta();
        }

        private void AddDestination(FolderNode folder)
        {
            if (folder == null || !_selectedKeys.Add(KeyOf(folder))) return;
            SelectedDestinations.Add(folder);
            NotifySelectionMeta();
        }

        private void OnRowSelectionChanged(SelectableFolder row)
        {
            if (row.IsSelected) AddDestination(row.Folder);
            else RemoveDestination(row.Folder);
        }

        private void NotifySelectionMeta()
        {
            OnPropertyChanged(nameof(HasSelectedDestinations));
            OnPropertyChanged(nameof(SelectionHeader));
            OnPropertyChanged(nameof(ClassifyButtonText));
        }

        /// <summary>Create a sub-folder under a specific folder (the per-row "+").</summary>
        public void CreateSubfolderUnder(FolderNode parent)
        {
            if (_isBusy || parent == null || _createSubfolder == null) return;

            string name = _promptForName != null ? _promptForName(parent.Name) : null;
            if (string.IsNullOrWhiteSpace(name)) return;

            var created = _createSubfolder(parent, name.Trim());
            if (created == null) { Status = _loc.GetString("Status_CouldNotCreateFolder"); return; }

            Status = _loc.GetString("Status_FolderCreated", created.Name, parent.Name);
            Refresh();
        }

        /// <summary>
        /// Runs the classify; true when it actually executed (not busy/empty/
        /// cancelled). Responsive (v2.4 D2): the moves stay on this STA thread
        /// but the pump is yielded between items, so Outlook repaints and stays
        /// interactive while a multi-item batch files, with per-item progress in
        /// the status line.
        /// </summary>
        private async Task<bool> DoClassifyAsync(IReadOnlyList<FolderNode> destinations)
        {
            if (_isBusy) return false; // ignore a repeat trigger while a classify is in flight

            // A pinned just-sent mail overrides the live selection for this one
            // classify (E1); otherwise file the current Outlook selection.
            var items = _pinnedItem != null
                ? new[] { _pinnedItem }
                : (_getSelection != null ? _getSelection() : new MailItemRef[0]);
            if (items.Count == 0) { Status = _loc.GetString("Status_NoMailSelectedAction"); return false; }

            IsBusy = true;
            Status = _loc.GetString("MainPane_Busy_Filing");

            try
            {
                var preflight = _classifier.Preflight(items, _widenConversation);

                bool markTasksComplete = false;
                if (preflight.FlaggedIncomplete.Count > 0 && _confirmMarkTasksComplete != null)
                {
                    var answer = _confirmMarkTasksComplete(preflight.FlaggedIncomplete.Count);
                    if (answer == null) { Status = _loc.GetString("Status_ClassifyCancelled"); return false; }
                    markTasksComplete = answer.Value;
                }

                bool safetyCopy = _settings != null
                    && _settings.GetBool(SettingsKeys.ClassifySafetyCopy, false);
                string bannerSignature = (_stripBanner && _settings != null)
                    ? _settings.Get(SettingsKeys.ExternalBannerSignature, null)
                    : null;

                // F2: in Modal mode, gather attachments and show the disposition
                // modal in the preflight (before the responsive move loop). A
                // cancel aborts the whole classify.
                IReadOnlyList<AttachmentDisposition> attachmentDispositions = null;
                if (_removeAttachments && _settings != null
                    && Settings.Load(_settings).AttachmentRemovalMode == AttachmentRemovalMode.Modal
                    && TaskPaneServices.GatherAttachments != null
                    && TaskPaneServices.ShowAttachmentDisposition != null)
                {
                    var groups = TaskPaneServices.GatherAttachments(preflight.Items);
                    if (groups.Any(g => g.IsEncrypted || g.Attachments.Count > 0))
                    {
                        attachmentDispositions = TaskPaneServices.ShowAttachmentDisposition(groups);
                        if (attachmentDispositions == null)
                        {
                            Status = _loc.GetString("Status_ClassifyCancelled");
                            return false;
                        }
                    }
                }

                // Core invokes Report synchronously on this (STA/UI) thread
                // between items, so set Status inline - the Background yield that
                // follows each item repaints it. (Progress<T> would Post the
                // callback to the ambient SynchronizationContext, which this host
                // does not set up, so the updates would not reach the binding.)
                var progress = new SynchronousProgress<ClassifyProgress>(p =>
                    Status = _loc.Plural(p.Completed, "Status_Classify_Progress_One",
                                         "Status_Classify_Progress_Other", p.Total));

                // F3: localized templates for the "former attachments" label,
                // applied to each filed copy whose attachments were disposed of -
                // unless the user chose to leave no trace.
                AttachmentLabelOptions labelOptions = null;
                if (attachmentDispositions != null && _settings != null
                    && Settings.Load(_settings).AttachmentLabelLocation != AttachmentLabelLocation.None)
                {
                    labelOptions = new AttachmentLabelOptions(
                        _loc.GetString("AttachmentLabel_Header_One"),
                        _loc.GetString("AttachmentLabel_Header_Other"),
                        _loc.GetString("AttachmentLabel_SavedTo"),
                        _loc.GetString("AttachmentLabel_DeletedOn"),
                        "yyyy-MM-dd");
                }

                var result = await _classifier.ClassifyAsync(
                    new ClassifyRequest(preflight.Items, destinations, _keepCopy,
                                        _removeAttachments, markTasksComplete, safetyCopy,
                                        bannerSignature, attachmentDispositions, labelOptions),
                    progress,
                    YieldToPump);

                string verb = _loc.GetString(_keepCopy ? "Status_Classify_Copied" : "Status_Classify_Filed");
                string failed = result.Errors > 0
                    ? _loc.GetString("Status_Classify_Failed", result.Errors)
                    : string.Empty;
                Status = _loc.Plural(result.ItemsProcessed, "Status_Classify_Result_One", "Status_Classify_Result_Other",
                                      verb, destinations.Count, failed);

                if (result.EncryptedStripSkips > 0)
                    Status += _loc.Plural(result.EncryptedStripSkips, "Status_Classify_EncryptedKept_One", "Status_Classify_EncryptedKept_Other");

                if (preflight.SkippedEncrypted.Count > 0)
                    Status += _loc.Plural(preflight.SkippedEncrypted.Count, "Status_Classify_EncryptedInConversation_One", "Status_Classify_EncryptedInConversation_Other");

                _lastUndo = result.Undo; // a non-undoable run clears the slot (null)
                OnPropertyChanged(nameof(CanUndo));

                ClearPin(); // E1: the pinned just-sent mail is now filed
                RefreshSelection();
                ClearQuerySilently();
                ResetOptionsToDefaults();
                return true;
            }
            finally
            {
                IsBusy = false;
            }
        }

        /// <summary>
        /// Release the STA message pump for one cycle at Background priority so
        /// Outlook repaints and processes input between classified items (v2.4
        /// D2). Resumes on this dispatcher, keeping the COM moves on the STA.
        /// </summary>
        private static async Task YieldToPump() =>
            await Dispatcher.Yield(DispatcherPriority.Background);

        /// <summary>
        /// An <see cref="IProgress{T}"/> that invokes the handler synchronously
        /// on the reporting thread, unlike <see cref="Progress{T}"/> which posts
        /// to a captured <see cref="System.Threading.SynchronizationContext"/>.
        /// The responsive classify reports on the UI thread, so this keeps the
        /// status update inline and ordered with the message-pump yields.
        /// </summary>
        private sealed class SynchronousProgress<T> : IProgress<T>
        {
            private readonly Action<T> _handler;
            public SynchronousProgress(Action<T> handler) => _handler = handler;
            public void Report(T value) => _handler(value);
        }

        /// <summary>
        /// Set the query to empty without triggering a re-search. Results remain
        /// visible until the user types the next query (v2.4).
        /// </summary>
        private void ClearQuerySilently()
        {
            _query = string.Empty;
            OnPropertyChanged(nameof(Query));
        }

        /// <summary>
        /// Reset the option checkboxes to their persisted settings defaults and
        /// collapse the Options panel, so the pane is clean for the next classify
        /// session (v2.4 B1).
        /// </summary>
        private void ResetOptionsToDefaults()
        {
            if (_settings != null)
            {
                // Set the backing fields (not the public setters): the AllResults
                // setter would call Refresh(), which clears the just-classified
                // results the user wants to keep seeing (v2.4 A1); the others
                // would re-write the same persisted value. Just notify the view.
                _allResults = _settings.GetBool(SettingsKeys.AllResults, false);
                _keepCopy = _settings.GetBool(SettingsKeys.KeepCopy, false);
                _removeAttachments = _settings.GetBool(SettingsKeys.RemoveAttachments, false);
                _widenConversation = _settings.GetBool(SettingsKeys.WidenConversation, false);
                _stripBanner = _settings.GetBool(SettingsKeys.StripBannerOnClassify, false);
                OnPropertyChanged(nameof(AllResults));
                OnPropertyChanged(nameof(KeepCopy));
                OnPropertyChanged(nameof(RemoveAttachments));
                OnPropertyChanged(nameof(WidenConversation));
                OnPropertyChanged(nameof(StripBanner));
            }
            IsOptionsExpanded = false;
        }

        /// <summary>
        /// Re-search after the typing pause configured in settings (immediate
        /// when the debounce is 0). Toggles and folder creation bypass this and
        /// call <see cref="Refresh"/> directly.
        /// </summary>
        private void ScheduleRefresh()
        {
            int debounceMs = _settings != null
                ? Settings.Load(_settings).SearchDebounceMs
                : Settings.DefaultSearchDebounceMs;
            if (debounceMs <= 0) { Refresh(); return; }

            if (_searchTimer == null)
            {
                _searchTimer = new DispatcherTimer(DispatcherPriority.Input);
                _searchTimer.Tick += (s, e) => { _searchTimer.Stop(); Refresh(); };
            }

            _searchTimer.Stop();
            _searchTimer.Interval = TimeSpan.FromMilliseconds(debounceMs);
            _searchTimer.Start();
        }

        /// <summary>Run a pending debounced search now (before acting on the results).</summary>
        private void FlushPendingSearch()
        {
            if (_searchTimer != null && _searchTimer.IsEnabled)
            {
                _searchTimer.Stop();
                Refresh();
            }
        }

        /// <summary>Add one result row, re-checking it if it is in the accumulated selection and tracking its checkbox.</summary>
        private void AddResultRow(FolderSearchResult r)
        {
            var row = new SelectableFolder(r);
            row.IsSelected = _selectedKeys.Contains(KeyOf(row.Folder));
            row.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(SelectableFolder.IsSelected))
                    OnRowSelectionChanged((SelectableFolder)s);
            };
            Results.Add(row);
        }

        private void Refresh()
        {
            var matchMode = FolderMatchMode.Substring;
            var maxResults = FolderSearchOptions.DefaultMaxResults;
            var minLength = FolderSearchOptions.DefaultMinQueryLength;
            if (_settings != null)
            {
                var settings = Settings.Load(_settings);
                matchMode = settings.FolderMatchMode;
                maxResults = settings.MaxResults;
                minLength = settings.MinSearchLength;
            }

            var outcome = _search.Search(_query,
                new FolderSearchOptions(matchMode, _allResults, maxResults, minLength));

            Results.Clear();
            foreach (var r in outcome.Results)
                AddResultRow(r);

            string trimmed = (_query ?? string.Empty).Trim();
            if (outcome.TotalMatchCount == 0)
            {
                if (trimmed.Length == 0)
                    Status = string.Empty;
                else if (trimmed.Length < minLength)
                    Status = _loc.GetString("Status_TypeAtLeastChars", minLength);
                else
                    Status = _loc.GetString("Status_NoMatchingFolders");
            }
            else if (outcome.LimitExceeded)
                Status = _loc.GetString("Status_ShowingResults", Results.Count, outcome.TotalMatchCount);
            else
                Status = _loc.Plural(outcome.TotalMatchCount, "Status_FoldersFound_One", "Status_FoldersFound_Other");
        }
    }
}
