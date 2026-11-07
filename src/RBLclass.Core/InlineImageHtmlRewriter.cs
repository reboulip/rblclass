using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace RBLclass.Core
{
    /// <summary>Outcome of <see cref="InlineImageHtmlRewriter.Rewrite"/>.</summary>
    public sealed class InlineImageRewriteResult
    {
        public InlineImageRewriteResult(string html, IReadOnlyCollection<string> replacedContentIds)
        {
            Html = html;
            ReplacedContentIds = replacedContentIds;
        }

        public string Html { get; }

        /// <summary>Content ids actually found and replaced; only those may be deleted.</summary>
        public IReadOnlyCollection<string> ReplacedContentIds { get; }
    }

    /// <summary>
    /// Replaces the body references to removed inline images with a short text
    /// placeholder, so no broken image remains (v2.11.0.0 C2). Pure: no Outlook
    /// types. Handles plain <c>&lt;img src="cid:..."&gt;</c> tags and the Word
    /// VML form Outlook writes (<c>&lt;!--[if gte vml 1]&gt;...&lt;v:imagedata
    /// src="cid:..."/&gt;...&lt;![endif]--&gt;&lt;![if !vml]&gt;&lt;img
    /// ...&gt;&lt;![endif]&gt;</c>), where Outlook renders the VML branch: the
    /// VML block is dropped and the fallback branch is unwrapped so the
    /// placeholder shows. Tags are found by a linear, quote-aware scan; a regex
    /// only ever runs on a single tag's text.
    /// </summary>
    public static class InlineImageHtmlRewriter
    {
        private const string VmlStart = "<!--[if gte vml 1]>";
        private const string VmlEnd = "<![endif]-->";
        private const string NotVmlStart = "<![if !vml]>";
        private const string DownlevelEnd = "<![endif]>";

        private static readonly Regex SrcAttribute = new Regex(
            @"\bsrc\s*=\s*(?:""(?<v>[^""]*)""|'(?<v>[^']*)'|(?<v>[^\s>]+))",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(250));

        public static InlineImageRewriteResult Rewrite(string html, IEnumerable<string> contentIds, string placeholder)
        {
            if (html == null) return new InlineImageRewriteResult(string.Empty, new string[0]);

            var targets = new HashSet<string>(
                (contentIds ?? Enumerable.Empty<string>())
                    .Select(InlineImageRules.NormalizeContentId)
                    .Where(id => id != null),
                StringComparer.OrdinalIgnoreCase);
            if (targets.Count == 0) return new InlineImageRewriteResult(html, new string[0]);

            var replaced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string text = DropVmlBlocks(html, targets, replaced);
            text = ReplaceImgTags(text, targets, replaced, WebUtility.HtmlEncode(placeholder ?? string.Empty));
            return new InlineImageRewriteResult(text, replaced.ToArray());
        }

        private static string DropVmlBlocks(string html, HashSet<string> targets, HashSet<string> replaced)
        {
            var sb = new StringBuilder(html.Length);
            int pos = 0;
            while (true)
            {
                int start = html.IndexOf(VmlStart, pos, StringComparison.OrdinalIgnoreCase);
                if (start < 0) break;
                int end = html.IndexOf(VmlEnd, start, StringComparison.OrdinalIgnoreCase);
                if (end < 0) break;
                end += VmlEnd.Length;

                string block = html.Substring(start, end - start);
                string id = FirstTargetIn(block, targets);
                if (id == null)
                {
                    sb.Append(html, pos, end - pos);
                    pos = end;
                    continue;
                }

                replaced.Add(id);
                sb.Append(html, pos, start - pos);
                pos = end;

                // Unwrap the fallback branch, which Outlook hides while VML renders.
                int next = SkipWhitespace(html, pos);
                if (string.Compare(html, next, NotVmlStart, 0, NotVmlStart.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    int close = html.IndexOf(DownlevelEnd, next + NotVmlStart.Length, StringComparison.OrdinalIgnoreCase);
                    if (close >= 0)
                    {
                        sb.Append(html, next + NotVmlStart.Length, close - (next + NotVmlStart.Length));
                        pos = close + DownlevelEnd.Length;
                    }
                }
            }
            sb.Append(html, pos, html.Length - pos);
            return sb.ToString();
        }

        private static string ReplaceImgTags(string html, HashSet<string> targets, HashSet<string> replaced, string placeholder)
        {
            var sb = new StringBuilder(html.Length);
            int pos = 0;
            while (true)
            {
                int start = IndexOfImgTag(html, pos);
                if (start < 0) break;
                int end = TagEnd(html, start);
                if (end < 0) break;

                string id = ContentIdOf(html.Substring(start, end - start));
                sb.Append(html, pos, start - pos);
                if (id != null && targets.Contains(id))
                {
                    replaced.Add(id);
                    sb.Append(placeholder);
                }
                else
                {
                    sb.Append(html, start, end - start);
                }
                pos = end;
            }
            sb.Append(html, pos, html.Length - pos);
            return sb.ToString();
        }

        private static int IndexOfImgTag(string html, int from)
        {
            int i = from;
            while (true)
            {
                i = html.IndexOf("<img", i, StringComparison.OrdinalIgnoreCase);
                if (i < 0) return -1;
                int after = i + 4;
                if (after >= html.Length || char.IsWhiteSpace(html[after]) || html[after] == '>' || html[after] == '/')
                    return i;
                i = after;
            }
        }

        /// <summary>Index just past the tag's closing '&gt;', ignoring any inside quoted attribute values.</summary>
        private static int TagEnd(string html, int start)
        {
            char quote = '\0';
            for (int i = start + 1; i < html.Length; i++)
            {
                char c = html[i];
                if (quote != '\0') { if (c == quote) quote = '\0'; }
                else if (c == '"' || c == '\'') quote = c;
                else if (c == '>') return i + 1;
            }
            return -1;
        }

        private static string ContentIdOf(string tag)
        {
            Match m;
            try { m = SrcAttribute.Match(tag); }
            catch (RegexMatchTimeoutException) { return null; }
            if (!m.Success) return null;

            string value = WebUtility.HtmlDecode(m.Groups["v"].Value).Trim();
            if (!value.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)) return null;
            try { value = Uri.UnescapeDataString(value.Substring(4)); }
            catch { value = value.Substring(4); }
            return InlineImageRules.NormalizeContentId(value);
        }

        private static string FirstTargetIn(string block, HashSet<string> targets)
        {
            int i = 0;
            while (true)
            {
                i = block.IndexOf("<", i, StringComparison.Ordinal);
                if (i < 0) return null;
                int end = TagEnd(block, i);
                if (end < 0) return null;
                string id = ContentIdOf(block.Substring(i, end - i));
                if (id != null && targets.Contains(id)) return id;
                i = end;
            }
        }

        private static int SkipWhitespace(string text, int i)
        {
            while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
            return i;
        }
    }
}
