using System.Collections.ObjectModel;
using System.Linq;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// One attachment row in the F2 disposition modal: delete it, or save it to a
    /// directory picked via the F1 favourite-folder keyword search (or the OS
    /// browse fallback).
    /// </summary>
    public sealed class AttachmentRowViewModel : ObservableObject
    {
        // Save-to is the default per user feedback (v2.7.1.0 A2): the default
        // disposition must not be destructive. Confirm stays gated until the
        // row gets a directory (or is switched to Delete/Keep).
        private AttachmentDispositionAction _action = AttachmentDispositionAction.SaveTo;
        private string _targetDirectory = string.Empty;
        private string _favoriteQuery = string.Empty;
        private bool _actionSetByUser;

        public AttachmentRowViewModel(MailItemRef item, AttachmentInfo info, string mailSubject)
        {
            Item = item;
            Info = info;
            MailSubject = mailSubject;
        }

        internal System.Action<string> DirectoryChosen;

        /// <summary>
        /// True once the user has explicitly picked an action for this row.
        /// Same-directory propagation must never override such a row (v2.7.1.0 A2).
        /// </summary>
        internal bool IsActionUserSet => _actionSetByUser;

        internal void SetDirectorySilent(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) return;
            TargetDirectory = directory;
            // Propagation, not a user gesture - must not mark the row user-set.
            SetActionCore(AttachmentDispositionAction.SaveTo);
            _favoriteQuery = string.Empty;
            OnPropertyChanged(nameof(FavoriteQuery));
            FavoriteResults.Clear();
        }

        public MailItemRef Item { get; }
        public AttachmentInfo Info { get; }
        public string MailSubject { get; }
        public string FileName => Info.FileName;
        public string SizeText => FormatSize(Info.SizeBytes);

        /// <summary>Keyword search over the favourite-folder index (F1).</summary>
        public ObservableCollection<string> FavoriteResults { get; } = new ObservableCollection<string>();

        public AttachmentDispositionAction Action
        {
            get => _action;
            set
            {
                _actionSetByUser = true;
                SetActionCore(value);
            }
        }

        private void SetActionCore(AttachmentDispositionAction action)
        {
            if (!SetProperty(ref _action, action)) return;
            OnPropertyChanged(nameof(IsDelete));
            OnPropertyChanged(nameof(IsSaveTo));
            OnPropertyChanged(nameof(IsKeep));
            // Leaving Save-to clears its directory search state so a stale
            // target can't gate Confirm or be re-applied (v2.5.0.0 B1).
            if (_action != AttachmentDispositionAction.SaveTo)
            {
                _favoriteQuery = string.Empty;
                OnPropertyChanged(nameof(FavoriteQuery));
                FavoriteResults.Clear();
            }
        }

        public bool IsDelete
        {
            get => _action == AttachmentDispositionAction.Delete;
            set { if (value) Action = AttachmentDispositionAction.Delete; }
        }

        public bool IsSaveTo
        {
            get => _action == AttachmentDispositionAction.SaveTo;
            set { if (value) Action = AttachmentDispositionAction.SaveTo; }
        }

        /// <summary>Leave this attachment on the filed copy untouched (v2.5.0.0 B1).</summary>
        public bool IsKeep
        {
            get => _action == AttachmentDispositionAction.Keep;
            set { if (value) Action = AttachmentDispositionAction.Keep; }
        }

        /// <summary>Directory the attachment will be saved into when <see cref="IsSaveTo"/>.</summary>
        public string TargetDirectory
        {
            get => _targetDirectory;
            set { if (SetProperty(ref _targetDirectory, value ?? string.Empty)) OnPropertyChanged(nameof(HasTarget)); }
        }

        public bool HasTarget => !string.IsNullOrWhiteSpace(_targetDirectory);

        /// <summary>Keyword query; updating it re-runs the favourite-folder search.</summary>
        public string FavoriteQuery
        {
            get => _favoriteQuery;
            set
            {
                if (!SetProperty(ref _favoriteQuery, value)) return;
                RefreshFavoriteResults();
            }
        }

        /// <summary>
        /// Choosing a search result (or browse target) sets the directory, flips
        /// to Save-to, and collapses the search list so only the chosen target
        /// remains visible.
        /// </summary>
        public void ChooseDirectory(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) return;
            TargetDirectory = directory;
            Action = AttachmentDispositionAction.SaveTo;
            _favoriteQuery = string.Empty;
            OnPropertyChanged(nameof(FavoriteQuery));
            FavoriteResults.Clear();
            DirectoryChosen?.Invoke(directory);
        }

        public void Browse()
        {
            var dir = TaskPaneServices.BrowseForFolder?.Invoke();
            ChooseDirectory(dir);
        }

        private void RefreshFavoriteResults()
        {
            FavoriteResults.Clear();
            var service = TaskPaneServices.FavoriteFolderService;
            if (service == null) return;
            var outcome = service.Search(_favoriteQuery);
            foreach (var path in outcome.Results.Select(r => r.Folder.FullPath))
                FavoriteResults.Add(path);
        }

        private static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024) + " KB";
            return (bytes / (1024 * 1024)) + " MB";
        }
    }
}
