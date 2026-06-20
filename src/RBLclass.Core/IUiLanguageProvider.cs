namespace RBLclass.Core
{
    /// <summary>
    /// Detects the host Outlook's UI language, so <see cref="UiLanguageResolver"/>
    /// can pick a matching RBLclass UI language when the user hasn't pinned one.
    /// </summary>
    public interface IUiLanguageProvider
    {
        /// <summary>
        /// Two-letter ISO 639-1 code for Outlook's UI language (e.g. "fr", "de",
        /// "en"), or null if it could not be determined.
        /// </summary>
        string GetOutlookUiLanguageCode();
    }
}
