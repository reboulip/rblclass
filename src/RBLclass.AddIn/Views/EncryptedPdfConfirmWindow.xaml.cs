using System.Windows;
using RBLclass.AddIn.Theming;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// Confirmation before exporting encrypted mails to PDF (v2.13.0.0 C1).
    /// DialogResult true = Continue, false = skip the encrypted mails (default, Esc, close).
    /// </summary>
    public partial class EncryptedPdfConfirmWindow : Window
    {
        public EncryptedPdfConfirmWindow()
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);
        }

        private void Continue_Click(object sender, RoutedEventArgs e) { DialogResult = true; Close(); }

        private void Skip_Click(object sender, RoutedEventArgs e) { DialogResult = false; Close(); }
    }
}
