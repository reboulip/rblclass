using System.Collections.ObjectModel;
using System.Linq;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// Directory-picker for the "Save as" action (v2.8.0.0 A1): reuses the same
    /// favourite-folder keyword-search + OS-browse UX as AttachmentRowViewModel,
    /// but for a single target directory (no per-attachment rows).
    /// </summary>
    public sealed class SaveAsPickerViewModel : ObservableObject
    {
        private string _favoriteQuery = string.Empty;
        private string _targetDirectory = string.Empty;
        private bool _askFormat;
        private bool _pdfSelected;

        public ObservableCollection<string> FavoriteResults { get; } =
            new ObservableCollection<string>();

        public string FavoriteQuery
        {
            get => _favoriteQuery;
            set { if (SetProperty(ref _favoriteQuery, value)) RefreshFavoriteResults(); }
        }

        public string TargetDirectory
        {
            get => _targetDirectory;
            set
            {
                if (SetProperty(ref _targetDirectory, value ?? string.Empty))
                    OnPropertyChanged(nameof(HasTarget));
            }
        }

        public bool HasTarget => !string.IsNullOrWhiteSpace(_targetDirectory);

        /// <summary>True when the format setting is "Ask each time": the picker shows the format choice.</summary>
        public bool AskFormat
        {
            get => _askFormat;
            set => SetProperty(ref _askFormat, value);
        }

        /// <summary>The picker's format answer: .msg is preselected, per action (never persisted).</summary>
        public bool IsPdfSelected
        {
            get => _pdfSelected;
            set { if (SetProperty(ref _pdfSelected, value)) OnPropertyChanged(nameof(IsMsgSelected)); }
        }

        public bool IsMsgSelected
        {
            get => !_pdfSelected;
            set { if (value) IsPdfSelected = false; }
        }

        public SaveAsFormat SelectedFormat => _pdfSelected ? SaveAsFormat.Pdf : SaveAsFormat.Msg;

        public void ChooseDirectory(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) return;
            TargetDirectory = directory;
            _favoriteQuery = string.Empty;
            OnPropertyChanged(nameof(FavoriteQuery));
            FavoriteResults.Clear();
        }

        public void Browse()
        {
            var dir = TaskPaneServices.BrowseForFolder?.Invoke();
            if (!string.IsNullOrWhiteSpace(dir)) ChooseDirectory(dir);
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
    }
}
