using System;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Threading;
using Serilog;

namespace RBLclass.AddIn.Localization
{
    /// <summary>
    /// <see cref="ILocalizationService"/> over <c>Resources/Strings.resx</c>.
    /// Constructed once at startup with the resolved language code; sets
    /// <see cref="Thread.CurrentUICulture"/>/<see cref="Thread.CurrentCulture"/>
    /// on the Outlook UI thread before any pane or window is created, so every
    /// <c>ResourceManager</c> lookup and format operation picks up the right
    /// culture without passing it around explicitly.
    /// </summary>
    public sealed class LocalizationService : ILocalizationService
    {
        private readonly ResourceManager _resources;
        private readonly CultureInfo _culture;

        public LocalizationService(string languageCode)
        {
            CurrentLanguage = RBLclass.Core.UiLanguageResolver.IsSupportedLanguage(languageCode)
                ? languageCode
                : "en";

            _culture = new CultureInfo(CultureNameFor(CurrentLanguage));
            Thread.CurrentThread.CurrentUICulture = _culture;
            Thread.CurrentThread.CurrentCulture = _culture;

            _resources = new ResourceManager("RBLclass.AddIn.Resources.Strings",
                Assembly.GetExecutingAssembly());
        }

        public string CurrentLanguage { get; }

        public string GetString(string key)
        {
            string value = _resources.GetString(key, _culture);
            if (value == null)
            {
                Log.Warning("Missing localization key '{Key}' for language '{Language}'", key, CurrentLanguage);
                return key;
            }
            return value;
        }

        public string GetString(string key, params object[] args) =>
            string.Format(_culture, GetString(key), args);

        public string Plural(int count, string oneKey, string otherKey, params object[] args)
        {
            string key = count == 1 ? oneKey : otherKey;

            var allArgs = new object[args.Length + 1];
            allArgs[0] = count;
            Array.Copy(args, 0, allArgs, 1, args.Length);

            return string.Format(_culture, GetString(key), allArgs);
        }

        private static string CultureNameFor(string languageCode)
        {
            switch (languageCode)
            {
                case "fr": return "fr-FR";
                case "de": return "de-DE";
                default: return "en-US";
            }
        }
    }
}
