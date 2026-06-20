namespace RBLclass.AddIn.ViewModels
{
    /// <summary>A UI language choice for the Settings language dropdown (value + display name).</summary>
    public sealed class UiLanguageOption
    {
        public UiLanguageOption(string code, string label)
        {
            Code = code;
            Label = label;
        }

        /// <summary>"Auto" | "en" | "fr" | "de" - stored in <see cref="RBLclass.Core.Settings.PreferredUiLanguage"/>.</summary>
        public string Code { get; }

        public string Label { get; }
    }
}
