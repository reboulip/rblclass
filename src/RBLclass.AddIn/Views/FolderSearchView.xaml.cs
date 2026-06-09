using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// The "Open a folder" WPF view, hosted in the Custom Task Pane via
    /// ElementHost. Thin code-behind: input gestures call view-model methods
    /// (no MVVM package - see RBLclass.AddIn.Mvvm).
    /// </summary>
    public partial class FolderSearchView : UserControl
    {
        public FolderSearchView()
        {
            InitializeComponent();
        }

        private FolderSearchViewModel Vm => DataContext as FolderSearchViewModel;

        private void QueryBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Vm?.OpenFirst();
                e.Handled = true;
            }
        }

        private void ResultsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Vm?.OpenSelected();
        }

        private void Open_Click(object sender, RoutedEventArgs e) => Vm?.OpenSelected();
    }
}
