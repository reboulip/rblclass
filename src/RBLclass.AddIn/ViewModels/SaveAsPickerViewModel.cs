using System.Collections.ObjectModel;
using System.Linq;
using RBLclass.AddIn.Mvvm;

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
