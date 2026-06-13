namespace RBLclass.AddIn.Localization
{
    /// <summary>
    /// Looks up UI strings for the language resolved once at startup (see
    /// <see cref="RBLclass.Core.UiLanguageResolver"/>). Backed by
    /// <c>Resources/Strings.resx</c> and its <c>fr</c>/<c>de</c> satellite
    /// resources.
    /// </summary>
    public interface ILocalizationService
    {
        /// <summary>Resolved UI language: "en", "fr" or "de".</summary>
        string CurrentLanguage { get; }

        /// <summary>Looks up <paramref name="key"/>, returning the key itself (and logging a warning) if missing.</summary>
        string GetString(string key);

        /// <summary>Looks up <paramref name="key"/> and formats it with <paramref name="args"/>.</summary>
        string GetString(string key, params object[] args);

        /// <summary>
        /// Picks <paramref name="oneKey"/> when <paramref name="count"/> is 1,
        /// otherwise <paramref name="otherKey"/>, and formats the result with
        /// <paramref name="count"/> as the first argument followed by
        /// <paramref name="args"/>.
        /// </summary>
        string Plural(int count, string oneKey, string otherKey, params object[] args);
    }
}
