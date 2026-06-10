using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Produces a leading-ellipsis truncation of a string so its END stays
    /// visible (e.g. "…/ Projects / Client ACME") when it is too wide for the
    /// space available. Folder paths in the narrow vertical task pane need the
    /// leaf and its nearest parents, not the store name at the front.
    /// <para>
    /// Pure logic: the width measurement is injected, so the WPF shell can pass
    /// a real text measurer while this stays framework-free and unit-testable.
    /// WPF itself has no native leading/path ellipsis (that is a Win32/WinForms
    /// feature), which is why this exists.
    /// </para>
    /// </summary>
    public static class PathEllipsis
    {
        /// <summary>Horizontal ellipsis (U+2026).</summary>
        public const string DefaultEllipsis = "…";

        /// <summary>
        /// Returns <paramref name="text"/> unchanged if it already fits within
        /// <paramref name="maxWidth"/>; otherwise the longest tail of it,
        /// prefixed with <paramref name="ellipsis"/>, that still fits. If not
        /// even one tail character fits alongside the ellipsis, returns just the
        /// ellipsis. <paramref name="measure"/> returns the rendered width of a
        /// candidate string.
        /// </summary>
        public static string TrimStart(string text, double maxWidth,
            Func<string, double> measure, string ellipsis = DefaultEllipsis)
        {
            if (measure == null) throw new ArgumentNullException(nameof(measure));
            if (string.IsNullOrEmpty(text)) return text ?? string.Empty;
            if (maxWidth <= 0) return text;
            if (measure(text) <= maxWidth) return text;

            // Drop characters from the front until the ellipsis-prefixed tail fits.
            for (int start = 1; start < text.Length; start++)
            {
                string candidate = ellipsis + text.Substring(start);
                if (measure(candidate) <= maxWidth)
                    return candidate;
            }
            return ellipsis;
        }
    }
}
