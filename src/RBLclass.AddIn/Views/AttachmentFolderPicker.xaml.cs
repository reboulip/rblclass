using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>Folder picker of the attachment disposition modal (v2.11.0.0 C1).</summary>
    public partial class AttachmentFolderPicker : UserControl
    {
        public AttachmentFolderPicker()
        {
            InitializeComponent();
        }

        private SaveAsPickerViewModel Picker => DataContext as SaveAsPickerViewModel;

        private void Browse_Click(object sender, RoutedEventArgs e) => Picker?.Browse();

        private void FavoriteResult_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ListBox)?.SelectedItem is string path) Picker?.ChooseDirectory(path);
        }

        private void TargetDirectory_Click(object sender, RoutedEventArgs e)
        {
            var directory = Picker?.TargetDirectory;
            if (string.IsNullOrWhiteSpace(directory)) return;
            try { Process.Start("explorer.exe", "\"" + directory + "\""); }
            catch { }
        }
    }
}
