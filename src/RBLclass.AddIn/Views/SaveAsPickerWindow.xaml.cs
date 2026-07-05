using System.Windows;
using System.Windows.Controls;
using RBLclass.AddIn.Theming;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// Directory-picker dialog for the "Save as" pane action (v2.8.0.0 A1).
    /// Confirm returns DialogResult true; Cancel/close returns false.
    /// </summary>
    public partial class SaveAsPickerWindow : Window
    {
        public SaveAsPickerWindow()
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);
        }

        private SaveAsPickerViewModel Vm => DataContext as SaveAsPickerViewModel;

        private void Save_Click(object sender, RoutedEventArgs e) { DialogResult = true; Close(); }

        private void Cancel_Click(object sender, RoutedEventArgs e) { DialogResult = false; Close(); }

        private void Browse_Click(object sender, RoutedEventArgs e) => Vm?.Browse();

        private void FavoriteResult_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var list = sender as ListBox;
            if (list?.SelectedItem is string path)
                Vm?.ChooseDirectory(path);
        }
    }
}
