using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using RBLclass.Core;
using Serilog;

namespace RBLclass.AddIn.Theming
{
    /// <summary>
    /// Themes the WPF task pane and dialogs to match Outlook's current look.
    /// Reads the Office UI theme and the Windows app theme from the registry,
    /// resolves a <see cref="ThemeMode"/> via the (unit-tested)
    /// <see cref="OfficeThemeResolver"/>, and merges the matching brush
    /// dictionary plus the shared implicit-style dictionary into a view/window
    /// root. Re-applying on each show picks up a theme the user changed while
    /// Outlook was running (there is no add-in-facing Office "theme changed"
    /// event to hook, so resolve-on-show is the pragmatic refresh point).
    /// </summary>
    internal static class ThemeService
    {
        private const string ControlsUri = "/RBLclass.AddIn;component/Views/Themes/Controls.xaml";
        private const string LightUri = "/RBLclass.AddIn;component/Views/Themes/Brushes.Light.xaml";
        private const string DarkUri = "/RBLclass.AddIn;component/Views/Themes/Brushes.Dark.xaml";

        /// <summary>Resolve the effective theme from the live registry signals.</summary>
        public static ThemeMode CurrentMode()
        {
            return OfficeThemeResolver.Resolve(ReadOfficeUiTheme(), ReadWindowsAppsUseLightTheme());
        }

        /// <summary>
        /// Merge the resolved theme into <paramref name="root"/> (a view or
        /// window), replacing any theme dictionaries applied earlier so a
        /// re-apply swaps cleanly. Never throws into the caller.
        /// </summary>
        public static ThemeMode Apply(FrameworkElement root)
        {
            var mode = CurrentMode();
            if (root == null) return mode;

            try
            {
                var dicts = root.Resources.MergedDictionaries;

                // Drop previously-applied theme dictionaries (idempotent re-apply).
                for (int i = dicts.Count - 1; i >= 0; i--)
                {
                    var src = dicts[i].Source;
                    if (src != null && src.OriginalString.IndexOf("/Themes/", StringComparison.OrdinalIgnoreCase) >= 0)
                        dicts.RemoveAt(i);
                }

                dicts.Add(Load(ControlsUri));
                dicts.Add(Load(mode == ThemeMode.Dark ? DarkUri : LightUri));

                // The root's own surface isn't covered by the implicit control
                // styles, so bind it to the themed brushes directly (DynamicResource
                // so it tracks a later re-apply).
                if (root is Control control)
                {
                    control.SetResourceReference(Control.BackgroundProperty, "Theme.Background");
                    control.SetResourceReference(Control.ForegroundProperty, "Theme.Foreground");
                }
            }
            catch (Exception ex)
            {
                try { Log.Warning(ex, "ThemeService.Apply failed; leaving default colours."); } catch { }
            }

            return mode;
        }

        /// <summary>WinForms host background matching the resolved theme (avoids a white flash around the WPF view).</summary>
        public static System.Drawing.Color WinFormsBackColor(ThemeMode mode)
        {
            return mode == ThemeMode.Dark
                ? System.Drawing.Color.FromArgb(0x2B, 0x2B, 0x2B)
                : System.Drawing.Color.White;
        }

        private static ResourceDictionary Load(string packUri)
        {
            return new ResourceDictionary { Source = new Uri(packUri, UriKind.Relative) };
        }

        private static int? ReadOfficeUiTheme()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Common"))
                {
                    var value = key?.GetValue("UI Theme");
                    if (value is int i) return i;
                }
            }
            catch (Exception ex) { try { Log.Debug(ex, "Reading Office UI Theme failed."); } catch { } }
            return null;
        }

        private static bool? ReadWindowsAppsUseLightTheme()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    var value = key?.GetValue("AppsUseLightTheme");
                    if (value is int i) return i != 0;
                }
            }
            catch (Exception ex) { try { Log.Debug(ex, "Reading Windows AppsUseLightTheme failed."); } catch { } }
            return null;
        }
    }
}
