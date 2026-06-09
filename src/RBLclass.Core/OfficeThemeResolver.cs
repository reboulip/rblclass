namespace RBLclass.Core
{
    /// <summary>
    /// Resolves the effective <see cref="ThemeMode"/> from the raw Office and
    /// Windows theme signals, so the WPF shell can colour the task pane and
    /// dialogs to match Outlook. Pure decision logic (no registry, no WPF) so it
    /// can be unit-tested in <c>RBLclass.Core</c>; the add-in reads the registry
    /// values and feeds them in.
    /// </summary>
    /// <remarks>
    /// The Office UI theme lives in
    /// <c>HKCU\Software\Microsoft\Office\16.0\Common\UI Theme</c> (REG_DWORD).
    /// Known explicit codes are mapped directly; everything else - including the
    /// "Use system setting" option and any future/unknown code - falls back to
    /// the Windows app theme. Keying unknown codes to Windows is deliberate: the
    /// explicit-code list drifts between Office builds, but the "follow Windows"
    /// fallback is always safe.
    /// <para>
    /// Colorful is a light theme (coloured title bar, light work area). Classic
    /// builds report it as 0; the dev machine's Current channel build reports it
    /// as 7. Both are mapped to Light. Older builds never emit 7 (their themes
    /// use 0/3/4/5), so claiming 7 for Colorful is safe across targets.
    /// </para>
    /// </remarks>
    public static class OfficeThemeResolver
    {
        // Office UI Theme codes that are unambiguously one mode regardless of the
        // Windows setting.
        private const int Colorful = 0;
        private const int ColorfulCurrentChannel = 7;
        private const int DarkGray = 3;
        private const int Black = 4;
        private const int White = 5;

        /// <summary>
        /// Resolve the effective theme. <paramref name="officeUiTheme"/> is the
        /// raw <c>UI Theme</c> DWORD (null if absent); <paramref name="windowsAppsUseLightTheme"/>
        /// is the Windows <c>AppsUseLightTheme</c> flag (true = light, false =
        /// dark, null if absent).
        /// </summary>
        public static ThemeMode Resolve(int? officeUiTheme, bool? windowsAppsUseLightTheme)
        {
            switch (officeUiTheme)
            {
                case Black:
                case DarkGray:
                    return ThemeMode.Dark;
                case White:
                case Colorful:
                case ColorfulCurrentChannel:
                    return ThemeMode.Light;
                default:
                    // "Use system setting" or an unknown/absent code: follow
                    // Windows, defaulting to Light when even that is unknown.
                    return windowsAppsUseLightTheme == false ? ThemeMode.Dark : ThemeMode.Light;
            }
        }
    }
}
