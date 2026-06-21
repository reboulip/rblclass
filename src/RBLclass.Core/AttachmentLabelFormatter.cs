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
                string line = d.Action == AttachmentDispositionAction.SaveTo
                    ? string.Format(options.SavedToTemplate, d.TargetDirectory)
                    : string.Format(options.DeletedOnTemplate, when.ToString(options.DateFormat));
                sb.Append("<div><b>").Append(Encode(d.FileName)).Append("</b> — ")
                  .Append(Encode(line)).Append("</div>");
            }
            sb.Append("</div>");
            return sb.ToString();
        }

        private static string Encode(string value) => WebUtility.HtmlEncode(value ?? string.Empty);
    }
}
