using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Resolves which of the supported UI languages RBLclass should use, given
    /// the user's preference (<see cref="Settings.PreferredUiLanguage"/>) and
    /// Outlook's own UI language (<see cref="IUiLanguageProvider"/>). Resolved
    /// once at add-in startup.
    /// </summary>
    public static class UiLanguageResolver
    {
        private static readonly string[] Supported = { "en", "fr", "de" };

        /// <summary>True if <paramref name="languageCode"/> is one of the languages RBLclass ships translations for.</summary>
        public static bool IsSupportedLanguage(string languageCode)
        {
            return languageCode != null && Array.IndexOf(Supported, languageCode) >= 0;
        }

        /// <summary>
        /// Picks the UI language: an explicit, supported <paramref name="preferredSetting"/>
        /// wins; otherwise a supported <paramref name="outlookLanguageCode"/>;
        /// otherwise English.
        /// </summary>
        public static string Resolve(string preferredSetting, string outlookLanguageCode)
        {
            if (IsSupportedLanguage(preferredSetting))
                return preferredSetting;

            if (IsSupportedLanguage(outlookLanguageCode))
                return outlookLanguageCode;

            return "en";
        }
    }
}
