using System.Windows;
using System.Windows.Input;
using RBLclass.AddIn.Theming;
using RBLclass.AddIn.ViewModels;
using RBLclass.Core;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// Small modal folder-picker (search + pick + OK/Cancel), used by the
    /// sent-item triage prompt's "Class…" action to choose a destination
    /// without disturbing the live Outlook selection that the Classify pane
    /// is built around. Sets <see cref="DialogResult"/> and <see cref="ChosenFolder"/>.
    /// </summary>
    public partial class FolderPickerWindow : Window
    {
        public FolderPickerWindow()
        {
            InitializeComponent();
            ThemeService.Apply(this);
            Loaded += (s, e) => QueryBox.Focus();
        }

        /// <summary>The chosen destination once <see cref="Window.DialogResult"/> is true.</summary>
        public FolderNode ChosenFolder { get; private set; }

        private FolderPickerViewModel Vm => DataContext as FolderPickerViewModel;

        private void QueryBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) Accept();
        }

        private void ResultsList_MouseDoubleClick(object sender, MouseButtonEventArgs e) => Accept();

        private void Ok_Click(object sender, RoutedEventArgs e) => Accept();

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Accept()
        {
            var folder = Vm?.ChosenFolder;
            if (folder == null) return;

            ChosenFolder = folder;
            DialogResult = true;
            Close();
        }
    }
}
