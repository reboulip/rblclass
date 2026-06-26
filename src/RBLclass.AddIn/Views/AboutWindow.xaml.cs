using System.Windows;
using RBLclass.AddIn.Theming;

namespace RBLclass.AddIn.Views
{
    public partial class AboutWindow : Window
    {
        public AboutWindow(System.Version version)
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);
            VersionText.Text = version.ToString();
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void EasterEgg_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            TaskPaneServices.ActivateEasterEgg?.Invoke();
            Close();
        }
    }
}
