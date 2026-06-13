using System;
using System.Globalization;
using RBLclass.Core;
using Serilog;
using Office = Microsoft.Office.Core;
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.Outlook.Adapter
{
    /// <summary>
    /// <see cref="IUiLanguageProvider"/> over the live Outlook Object Model.
    /// Reads Outlook's UI display language so <see cref="UiLanguageResolver"/>
    /// can pick a matching RBLclass language when the user hasn't pinned one.
    /// </summary>
    /// <remarks>
    /// Touches COM and MUST be called on the Outlook UI (STA) thread, once at
    /// startup. Excluded by design from automated tests (needs a live Outlook).
    /// </remarks>
    public sealed class OutlookUiLanguageProvider : IUiLanguageProvider
    {
        private readonly OutlookOM.Application _app;

        public OutlookUiLanguageProvider(OutlookOM.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public string GetOutlookUiLanguageCode()
        {
            try
            {
                using (var settings = new ComRef<Office.LanguageSettings>(_app.LanguageSettings))
                {
                    int lcid = settings.Value.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
                    return new CultureInfo(lcid).TwoLetterISOLanguageName;
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Could not determine Outlook's UI language; defaulting to English");
                return null;
            }
        }
    }
}
