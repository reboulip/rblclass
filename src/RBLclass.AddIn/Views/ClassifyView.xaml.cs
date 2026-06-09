using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// The Classify WPF view. Destinations are toggled via per-row checkboxes;
    /// the view model gathers the checked folders. Enter files into the first
    /// result; the Classify button files into all checked folders.
    /// </summary>
    public partial class ClassifyView : UserControl
    {
        public ClassifyView()
        {
            InitializeComponent();
        }

        private ClassifyViewModel Vm => DataContext as ClassifyViewModel;

        private void QueryBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Vm?.ClassifyToFirst();
                e.Handled = true;
            }
        }

        private void Classify_Click(object sender, RoutedEventArgs e) => Vm?.ClassifyChecked();

        private void NewSubfolder_Click(object sender, RoutedEventArgs e)
        {
            var vm = Vm;
            if (vm == null) return;

            vm.CreateSubfolder(NewFolderNameBox.Text);
            NewFolderNameBox.Clear();
        }
    }
}
