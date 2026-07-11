using System;
using System.Collections.Generic;
using System.Net;
using System.Text;

namespace RBLclass.Core
{
    /// <summary>
    /// Composes the HTML "former attachments" note block recorded on a filed
    /// mail after disposition (v2.4.0.0 F3): one line per attachment with where
    /// it went ("Saved to …") or that it was deleted ("Deleted on …"). Pure -
    /// all localized text is supplied via <see cref="AttachmentLabelOptions"/> so
    /// the formatter is UI-free and unit-testable.
    /// </summary>
    public static class AttachmentLabelFormatter
    {
        public static string Format(IReadOnlyList<AttachmentDisposition> dispositions,
                                    AttachmentLabelOptions options, DateTime when)
        {
            if (dispositions == null || dispositions.Count == 0 || options == null)
                return null;

            string header = dispositions.Count == 1 ? options.HeaderOne : options.HeaderOther;

            var sb = new StringBuilder();
            sb.Append("<div style=\"border-left:3px solid #888;margin:8px 0;padding:4px 8px;font-size:0.9em;\">");
            sb.Append("<strong>").Append(Encode(header)).Append("</strong>");
            foreach (var d in dispositions)
            {
                string line;
                if (d.Action == AttachmentDispositionAction.SaveTo)
                {
                    // The saved-to directory becomes a file:/// hyperlink to the
                    // containing folder (v2.9.0.0 D1) - an HTML mail body cannot
                    // invoke explorer-select, so the folder link is the best
                    // available. The localized template is split on {0} so only
                    // the path is wrapped in the anchor, and each fragment is
                    // HTML-encoded separately.
                    var parts = options.SavedToTemplate.Split(
                        new[] { "{0}" }, 2, StringSplitOptions.None);
                    string pre = Encode(parts[0]);
                    string post = parts.Length > 1 ? Encode(parts[1]) : string.Empty;
                    line = pre
                         + "<a href=\"" + PathToFileUri(d.TargetDirectory) + "\">"
                         + Encode(d.TargetDirectory)
                         + "</a>"
                         + post;
                }
                else
                {
                    line = Encode(string.Format(
                        options.DeletedOnTemplate, when.ToString(options.DateFormat)));
                }
                sb.Append("<div><b>").Append(Encode(d.FileName)).Append("</b> — ")
                  .Append(line).Append("</div>");
            }
            sb.Append("</div>");
            return sb.ToString();
        }

        private static string Encode(string value) => WebUtility.HtmlEncode(value ?? string.Empty);

        /// <summary>
        /// A local Windows path as a file:/// URI for an href attribute
        /// (percent-encodes spaces and non-ASCII). Falls back to a plain
        /// slash-swap when <see cref="Uri"/> cannot parse the path.
        /// </summary>
        private static string PathToFileUri(string localPath)
        {
            if (string.IsNullOrEmpty(localPath)) return "#";
            try { return new Uri(localPath).AbsoluteUri; }
            catch { return "file:///" + localPath.Replace('\\', '/'); }
        }
    }
}
