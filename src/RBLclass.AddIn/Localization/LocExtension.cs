using System;
using System.Windows.Markup;

namespace RBLclass.AddIn.Localization
{
    /// <summary>
    /// XAML markup extension for localized strings: <c>{loc:Loc Key=MainPane_Undo_Button}</c>.
    /// Resolves once at XAML-load time via <see cref="TaskPaneServices.Localization"/>,
    /// which is published by the composition root before any view is constructed.
    /// </summary>
    public sealed class LocExtension : MarkupExtension
    {
        public LocExtension()
        {
        }

        public LocExtension(string key)
        {
            Key = key;
        }

        public string Key { get; set; }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            var localization = TaskPaneServices.Localization;
            return localization != null ? localization.GetString(Key) : Key;
        }
    }
}
