namespace RBLclass.Core
{
    /// <summary>
    /// Behaviour of the classify "remove attachments" option (v2.4.0.0 F2).
    /// </summary>
    public enum AttachmentRemovalMode
    {
        /// <summary>Show the per-attachment disposition modal before stripping (default).</summary>
        Modal = 0,

        /// <summary>Silently strip all attachments, the pre-F2 behaviour.</summary>
        DeleteSilently = 1
    }
}
