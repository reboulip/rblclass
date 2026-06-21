namespace RBLclass.Core
{
    /// <summary>
    /// Canonical keys for the <see cref="ISettingsStore"/>, mapping to the
    /// legacy options (minus the dropped QuickOpen) plus the settings
    /// introduced during the rewrite. All surfaced together by the Step 9
    /// settings dialog (see <see cref="Settings"/>); several were already
    /// wired into individual panes before Step 9 added the typed snapshot.
    /// </summary>
    public static class SettingsKeys
    {
        /// <summary>Open folders in a new explorer window (legacy settingsNewWindow).</summary>
        public const string OpenInNewWindow = "OpenInNewWindow";

        /// <summary>Expand all matching folders vs collapse (legacy settingsAllResults).</summary>
        public const string AllResults = "AllResults";

        /// <summary>Folder match mode: "WordPrefix" or "Substring".</summary>
        public const string FolderMatchMode = "FolderMatchMode";

        /// <summary>Cap on displayed folder-search results (legacy settingsMaxResults).</summary>
        public const string MaxResults = "MaxResults";

        /// <summary>Minimum (trimmed) query length before folder search runs (v2.2).</summary>
        public const string MinSearchLength = "MinSearchLength";

        /// <summary>
        /// The learned external-sender banner block (verbatim HTML), captured
        /// once from a sample mail; empty when the user hasn't taught one (v2.2).
        /// </summary>
        public const string ExternalBannerSignature = "ExternalBannerSignature";

        /// <summary>Auto-strip the learned banner from reply/forward drafts (v2.2, default off).</summary>
        public const string StripBannerOnReply = "StripBannerOnReply";

        /// <summary>Default state of the classify-time "strip banner from the filed copy" tickbox (v2.2).</summary>
        public const string StripBannerOnClassify = "StripBannerOnClassify";

        /// <summary>Pause after the last keystroke before folder search fires, in ms (v2.2).</summary>
        public const string SearchDebounceMs = "SearchDebounceMs";

        /// <summary>Keep originals after classify, i.e. don't delete (inverse of legacy settingsdeleteProcessedElts).</summary>
        public const string KeepCopy = "KeepCopy";

        /// <summary>Strip attachments when classifying (legacy settingsRemoveAttachments).</summary>
        public const string RemoveAttachments = "RemoveAttachments";

        /// <summary>
        /// Also place a copy of each classified mail in Deleted Items when not
        /// keeping the original (v2.2). Off by default: a move never loses data
        /// and Undo is the designed guardrail; this restores the old
        /// delete-after-copy side effect for users who relied on it.
        /// </summary>
        public const string ClassifySafetyCopy = "ClassifySafetyCopy";

        /// <summary>Widen the classify set to conversation siblings in Inbox/Sent (legacy settingsWholeConversation).</summary>
        public const string WidenConversation = "WidenConversation";

        /// <summary>Warn before sending to recipients outside the org (legacy settingsSendExt).</summary>
        public const string SendExternalWarning = "SendExternalWarning";

        /// <summary>
        /// Semicolon-separated domains treated as internal in addition to
        /// Exchange-resolved recipients (empty by default - nothing
        /// org-specific is hard-coded; Step 9 exposes this for editing).
        /// </summary>
        public const string InternalDomains = "InternalDomains";

        /// <summary>
        /// Semicolon-separated keywords that trigger the forgotten-attachment
        /// warning when none of them are backed by an actual attachment
        /// (legacy hard-coded list: attach/enclos/joint/PJ). No on/off toggle -
        /// the legacy guard had none either.
        /// </summary>
        public const string ForgottenAttachmentKeywords = "ForgottenAttachmentKeywords";

        /// <summary>
        /// What to do with a mail that lands in Sent Items: one of
        /// <see cref="SentItemTriageMode"/> (stored by name). Supersedes
        /// <see cref="SentItemTriagePrompt"/>.
        /// </summary>
        public const string SentItemTriageMode = "SentItemTriageMode";

        /// <summary>
        /// Legacy on/off sent-item triage flag (legacy settingsUFsent). Kept only
        /// to migrate existing installs to <see cref="SentItemTriageMode"/>; no
        /// longer written.
        /// </summary>
        public const string SentItemTriagePrompt = "SentItemTriagePrompt";

        /// <summary>
        /// Preferred UI language: "Auto" (follow Outlook's UI language, falling
        /// back to English), or one of the supported two-letter codes ("en",
        /// "fr", "de"). Resolved once at add-in startup by
        /// <see cref="UiLanguageResolver"/>.
        /// </summary>
        public const string PreferredUiLanguage = "PreferredUiLanguage";

        /// <summary>
        /// Last dock position of the task pane, stored as the integer value of
        /// <c>Microsoft.Office.Core.MsoCTPDockPosition</c>
        /// (msoCTPDockPositionRight = 2 is the default).
        /// </summary>
        public const string PaneDockPosition = "PaneDockPosition";

        /// <summary>
        /// After sent-item triage moves a mail to the Inbox, reveal the classify
        /// pane with that mail pinned as the next target so the user can file it
        /// immediately (v2.4.0.0 E1, default on).
        /// </summary>
        public const string ClassifyAfterMoveToInbox = "ClassifyAfterMoveToInbox";
    }
}
