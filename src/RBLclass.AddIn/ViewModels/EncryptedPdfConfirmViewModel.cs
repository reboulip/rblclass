namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// Text for the one-per-batch confirmation shown before a PDF export that
    /// includes encrypted mails (v2.13.0.0 C1): the PDF would contain them unencrypted.
    /// </summary>
    public sealed class EncryptedPdfConfirmViewModel
    {
        public EncryptedPdfConfirmViewModel(int encryptedCount)
        {
            Body = TaskPaneServices.Localization.Plural(encryptedCount,
                "EncryptedPdf_Body_One", "EncryptedPdf_Body_Other");
        }

        public string Body { get; }
    }
}
