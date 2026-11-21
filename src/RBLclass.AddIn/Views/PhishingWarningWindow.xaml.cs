using System.Windows;
using RBLclass.AddIn.Theming;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// The "Possible phishing" modal shared by the attachment and link warnings
    /// (v2.13.0.0 D2). Cancel is the default button and Esc; closing the window also
    /// cancels. <see cref="OpenAnyway"/> is true only after an explicit click.
    /// </summary>
    public partial class PhishingWarningWindow : Window
    {
        public PhishingWarningWindow()
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);
        }

        /// <summary>True when the user chose "Open anyway".</summary>
        public bool OpenAnyway { get; private set; }

        /// <summary>Show the warning modally; returns true for "Open anyway", false for Cancel.</summary>
        public static bool Confirm(PhishingWarningViewModel viewModel)
        {
            var window = new PhishingWarningWindow { DataContext = viewModel };
            window.ShowDialog();
            return window.OpenAnyway;
        }

        private void OpenAnyway_Click(object sender, RoutedEventArgs e)
        {
            OpenAnyway = true;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            OpenAnyway = false;
            DialogResult = false;
            Close();
        }
    }
}
