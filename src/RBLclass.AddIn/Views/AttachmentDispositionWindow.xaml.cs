using System.Windows;
using System.Windows.Controls;
using RBLclass.AddIn.Theming;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// The F2 per-attachment disposition modal (v2.4.0.0). Confirm returns
    /// DialogResult true; Cancel/close returns false, which aborts the classify.
    /// </summary>
    public partial class AttachmentDispositionWindow : Window
    {
        public AttachmentDispositionWindow()
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);
        }

        private AttachmentDispositionViewModel Vm => DataContext as AttachmentDispositionViewModel;

        private void Confirm_Click(object sender, RoutedEventArgs e) { DialogResult = true; Close(); }

        private void Cancel_Click(object sender, RoutedEventArgs e) { DialogResult = false; Close(); }

        private void DeleteAll_Click(object sender, RoutedEventArgs e) => Vm?.DeleteAll();

        private void Browse_Click(object sender, RoutedEventArgs e) =>
            ((sender as FrameworkElement)?.DataContext as AttachmentRowViewModel)?.Browse();

        private void FavoriteResult_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var list = sender as ListBox;
            if (list?.SelectedItem is string path && list.DataContext is AttachmentRowViewModel row)
                row.ChooseDirectory(path);
        }
    }
}
