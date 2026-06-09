namespace RBLclass.Core
{
    /// <summary>
    /// The light/dark axis the UI themes against. This is the only distinction
    /// that matters for readability (the pilot reported dark-on-dark text and a
    /// white pane under a dark Outlook); accent nuances (Colorful vs White) are
    /// not modelled here.
    /// </summary>
    public enum ThemeMode
    {
        Light = 0,
        Dark = 1
    }
}
