namespace RBLclass.Core
{
    /// <summary>
    /// Where the "former attachments" label is recorded after disposition
    /// (v2.4.0.0 F3). Body is the reliable, durable path; InfoBar needs an
    /// Outlook form region and is deferred - treated as Body until implemented.
    /// </summary>
    public enum AttachmentLabelLocation
    {
        Body = 0,
        InfoBar = 1,

        /// <summary>Leave no trace - do not record a former-attachments note.</summary>
        None = 2
    }
}
