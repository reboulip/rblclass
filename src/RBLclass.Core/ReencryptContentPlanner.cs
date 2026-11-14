using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace RBLclass.Core
{
    /// <summary>
    /// Applies the classify actions to a decrypted <see cref="MailContent"/> for
    /// re-encrypted filing (v2.12.0.0 A3): the same rules the standard path
    /// applies to a plain filed mail (<c>ClassifierService.ApplyAttachmentRemoval</c>
    /// then the banner strip), on content instead of on an Outlook item.
    /// Pure: the input is never changed, surviving parts are reused as-is.
    /// </summary>
    public static class ReencryptContentPlanner
    {
        /// <param name="removeAttachments">The "remove attachments" option (false after a failed "Save to").</param>
        /// <param name="itemRows">This item's disposition rows only (the caller filters by item).</param>
        public static ReencryptContentPlan Plan(MailContent content, bool removeAttachments,
                                                IReadOnlyList<AttachmentDisposition> itemRows,
                                                string bannerSignature, AttachmentLabelOptions labelOptions,
                                                DateTime now)
        {
            if (content == null) throw new ArgumentNullException(nameof(content));
            var rows = itemRows ?? new AttachmentDisposition[0];

            var parts = content.Parts.ToList();
            string html = content.HtmlBody;
            string plain = content.PlainBody;
            int removed = 0;
            bool mismatch = false;
            var labelRows = new List<AttachmentDisposition>();

            if (removeAttachments || rows.Any(r => r.IsInlineImage))
            {
                var regular = rows.Where(r => !r.IsInlineImage).ToList();
                var inline = rows.Where(r => r.IsInlineImage).ToList();

                bool stripped = false;
                if (removeAttachments && (rows.Count == 0 || regular.Count > 0))
                {
                    if (regular.Any(r => r.Action == AttachmentDispositionAction.Keep))
                    {
                        // Per-id removal: never trust position. A row whose part is
                        // missing or renamed means the view changed since the modal.
                        foreach (var row in regular.Where(r => r.Action != AttachmentDispositionAction.Keep))
                        {
                            var part = parts.FirstOrDefault(p => p.AttachmentId == row.AttachmentId);
                            if (part == null || !string.Equals(part.FileName, row.FileName, StringComparison.OrdinalIgnoreCase))
                            {
                                mismatch = true;
                                continue;
                            }
                            parts.Remove(part);
                            removed++;
                            stripped = true;
                        }
                    }
                    else
                    {
                        // Strip all: every non-inline part goes (as the adapter's
                        // RemoveAttachments, "stripped" even when there was none).
                        removed += parts.RemoveAll(p => !p.IsInline);
                        stripped = true;
                    }
                }

                if (stripped)
                    labelRows.AddRange(regular.Where(r => r.Action != AttachmentDispositionAction.Keep));

                // Body images: removed by rewriting the HTML; only the parts whose
                // reference was actually replaced are dropped and labelled.
                var kept = new HashSet<string>(
                    inline.Where(r => r.Action == AttachmentDispositionAction.Keep).Select(r => r.ContentId),
                    StringComparer.OrdinalIgnoreCase);
                var removedRows = inline.Where(r => r.Action != AttachmentDispositionAction.Keep
                                                    && !kept.Contains(r.ContentId)).ToList();
                if (removedRows.Count > 0 && !string.IsNullOrEmpty(html))
                {
                    var rewrite = InlineImageHtmlRewriter.Rewrite(
                        html, removedRows.Select(r => r.ContentId).Distinct(StringComparer.OrdinalIgnoreCase),
                        removedRows[0].InlinePlaceholder);
                    var replaced = new HashSet<string>(rewrite.ReplacedContentIds, StringComparer.OrdinalIgnoreCase);
                    if (replaced.Count > 0)
                    {
                        html = rewrite.Html;
                        removed += parts.RemoveAll(p => p.ContentId != null && replaced.Contains(p.ContentId));
                        labelRows.AddRange(removedRows.Where(r => replaced.Contains(r.ContentId)));
                    }
                }
            }

            // Label first, banner after - the standard path's order.
            if (labelOptions != null && labelRows.Count > 0)
            {
                string block = AttachmentLabelFormatter.Format(labelRows, labelOptions, now);
                if (block != null)
                {
                    if (string.IsNullOrEmpty(html))
                    {
                        // A plain mail gets an HTML body, as assigning HTMLBody does
                        // on the standard path; the plain body is kept alongside.
                        html = "<html><body>" + block + "<div style=\"white-space:pre-wrap\">"
                             + WebUtility.HtmlEncode(plain ?? string.Empty) + "</div></body></html>";
                    }
                    else
                    {
                        html = AttachmentLabelFormatter.InsertAfterBodyTag(html, block);
                    }
                }
            }

            bool bannerStripped = false;
            if (!string.IsNullOrWhiteSpace(bannerSignature) && !string.IsNullOrEmpty(html))
            {
                string updated = ExternalBannerStripper.Strip(html, bannerSignature, out bannerStripped);
                if (bannerStripped) html = updated;
            }

            return new ReencryptContentPlan(new MailContent(html, plain, parts), labelRows,
                                            removed, bannerStripped, mismatch);
        }
    }

    /// <summary>The outcome of <see cref="ReencryptContentPlanner.Plan"/>.</summary>
    public sealed class ReencryptContentPlan
    {
        public ReencryptContentPlan(MailContent content, IReadOnlyList<AttachmentDisposition> labelledRows,
                                    int removedPartCount, bool bannerStripped, bool attachmentMapMismatch)
        {
            Content = content;
            LabelledRows = labelledRows ?? new AttachmentDisposition[0];
            RemovedPartCount = removedPartCount;
            BannerStripped = bannerStripped;
            AttachmentMapMismatch = attachmentMapMismatch;
        }

        /// <summary>The content to encrypt.</summary>
        public MailContent Content { get; }

        /// <summary>The rows recorded in the "former attachments" label.</summary>
        public IReadOnlyList<AttachmentDisposition> LabelledRows { get; }

        public int RemovedPartCount { get; }

        public bool BannerStripped { get; }

        /// <summary>
        /// A row to remove had no matching part (id and file name): the in-memory
        /// view changed since the attachment window. The caller must not file this
        /// content and falls back to the standard path.
        /// </summary>
        public bool AttachmentMapMismatch { get; }
    }
}
