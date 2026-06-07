namespace RBLclass.Core
{
    /// <summary>
    /// Canonical keys for the <see cref="ISettingsStore"/>, mapping to the nine
    /// legacy options. Only the ones with UI today are read/written; Step 9
    /// surfaces the rest in a settings pane.
    /// </summary>
    public static class SettingsKeys
    {
        /// <summary>Open folders in a new explorer window (legacy settingsNewWindow).</summary>
        public const string OpenInNewWindow = "OpenInNewWindow";

        /// <summary>Expand all matching folders vs collapse (legacy settingsAllResults).</summary>
        public const string AllResults = "AllResults";

        /// <summary>Folder match mode: "WordPrefix" or "Substring".</summary>
        public const string FolderMatchMode = "FolderMatchMode";

        /// <summary>Keep originals after classify, i.e. don't delete (inverse of legacy settingsdeleteProcessedElts).</summary>
        public const string KeepCopy = "KeepCopy";

        /// <summary>Strip attachments when classifying (legacy settingsRemoveAttachments).</summary>
        public const string RemoveAttachments = "RemoveAttachments";

        /// <summary>Widen the classify set to conversation siblings in Inbox/Sent (legacy settingsWholeConversation).</summary>
        public const string WidenConversation = "WidenConversation";
    }
}
