namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// A choice in the Settings re-encryption certificate list (v2.12.0.0 A1):
    /// thumbprint (empty = Automatic) and display label.
    /// </summary>
    public sealed class CertificateOption
    {
        public CertificateOption(string thumbprint, string label)
        {
            Thumbprint = thumbprint ?? string.Empty;
            Label = label;
        }

        /// <summary>Stored in <see cref="RBLclass.Core.Settings.ReencryptCertificateThumbprint"/>.</summary>
        public string Thumbprint { get; }

        public string Label { get; }
    }
}
