using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace RBLclass.Core
{
    public enum BannerDiagnosticOutcome { NoSignature, NoSelection, Found, NotFound }

    public struct BannerDiagnosticResult
    {
        public BannerDiagnosticOutcome Outcome { get; }
        public int SignatureLength { get; }
        public bool MatchedExact { get; }
        public BannerDiagnosticResult(BannerDiagnosticOutcome outcome, int sigLen, bool exact)
        { Outcome = outcome; SignatureLength = sigLen; MatchedExact = exact; }
    }

    /// <summary>
    /// Learns the "external sender" reminder banner from one captured sample and
    /// strips it back out of other mail bodies (v2.2). The banner is the block a
    /// mail-transport rule prepends to inbound external mail; it is the same HTML
    /// every time, so we store the captured block verbatim ("exact HTML block
    /// match") and remove an occurrence of it from a target body.
    /// </summary>
    /// <remarks>
    /// Pure string logic, kept in Core so it is unit-testable - the Outlook
    /// adapter only reads/writes the HTML body. Matching is exact, with the one
    /// concession that runs of whitespace are treated as equivalent (Outlook
    /// reflows insignificant whitespace in <c>HTMLBody</c> between mails), so a
    /// banner that is structurally identical but rewrapped still matches.
    /// </remarks>
    public static class ExternalBannerStripper
    {
        // Block-level tags a banner is built from. Transport-rule banners are
        // almost always a single coloured <table>; <div>/<p>/<blockquote> are
        // fallbacks for the rarer plain variants.
        private static readonly string[] BlockTags = { "table", "div", "blockquote", "p" };

        private static readonly TimeSpan MatchTimeout = TimeSpan.FromMilliseconds(250);

        /// <summary>
        /// Extract the banner block to remember from a captured mail body: the
        /// first block-level element (a <c>&lt;table&gt;</c> by preference, else
        /// the earliest <c>&lt;div&gt;</c>/<c>&lt;blockquote&gt;</c>/<c>&lt;p&gt;</c>),
        /// returned verbatim including its tags. Null when no block is found.
        /// The caller stores this as the signature.
        /// </summary>
        public static string ExtractBannerBlock(string htmlBody)
        {
            if (string.IsNullOrEmpty(htmlBody)) return null;

            // Start after <body> if present, so the document wrapper isn't picked.
            int from = 0;
            var bodyOpen = Regex.Match(htmlBody, "<body\\b[^>]*>",
                                       RegexOptions.IgnoreCase, MatchTimeout);
            if (bodyOpen.Success) from = bodyOpen.Index + bodyOpen.Length;

            // Prefer the first <table>; fall back to the earliest other block.
            var block = ExtractFirstOfTag(htmlBody, from, "table");
            if (block != null) return block;

            string earliest = null;
            int earliestIdx = int.MaxValue;
            foreach (var tag in BlockTags)
            {
                if (tag == "table") continue;
                int idx = IndexOfOpenTag(htmlBody, from, tag);
                if (idx >= 0 && idx < earliestIdx)
                {
                    var candidate = ExtractFirstOfTag(htmlBody, from, tag);
                    if (candidate != null) { earliest = candidate; earliestIdx = idx; }
                }
            }
            return earliest;
        }

        /// <summary>
        /// Remove the first occurrence of <paramref name="bannerBlock"/> from
        /// <paramref name="htmlBody"/>. Tries an exact match first, then a
        /// whitespace-tolerant match. Returns the (possibly unchanged) body and
        /// reports via <paramref name="stripped"/> whether anything was removed.
        /// </summary>
        public static string Strip(string htmlBody, string bannerBlock, out bool stripped)
        {
            stripped = false;
            if (string.IsNullOrEmpty(htmlBody) || string.IsNullOrWhiteSpace(bannerBlock))
                return htmlBody;

            string needle = bannerBlock.Trim();

            // Fast path: the transport rule inserts identical HTML, so an exact
            // ordinal match is the common case.
            int idx = htmlBody.IndexOf(needle, StringComparison.Ordinal);
            if (idx >= 0)
            {
                stripped = true;
                return Collapse(htmlBody.Remove(idx, needle.Length));
            }

            // Tolerant path: same non-whitespace structure, any whitespace runs.
            var pattern = BuildWhitespaceTolerantPattern(needle);
            if (pattern != null)
            {
                try
                {
                    var m = Regex.Match(htmlBody, pattern, RegexOptions.None, MatchTimeout);
                    if (m.Success)
                    {
                        stripped = true;
                        return Collapse(htmlBody.Remove(m.Index, m.Length));
                    }
                }
                catch (RegexMatchTimeoutException)
                {
                    // Give up tolerantly - leave the body untouched rather than risk a hang.
                }
            }

            return htmlBody;
        }

        /// <summary>Convenience overload that just says whether the banner is present/removable.</summary>
        public static bool ContainsBanner(string htmlBody, string bannerBlock)
        {
            bool stripped;
            Strip(htmlBody, bannerBlock, out stripped);
            return stripped;
        }

        /// <summary>
        /// Run the detection pipeline against <paramref name="htmlBody"/> and
        /// return a structured result for the Settings diagnostics UI (B3).
        /// </summary>
        public static BannerDiagnosticResult Diagnose(string htmlBody, string bannerSignature)
        {
            if (string.IsNullOrWhiteSpace(bannerSignature))
                return new BannerDiagnosticResult(BannerDiagnosticOutcome.NoSignature, 0, false);
            if (string.IsNullOrEmpty(htmlBody))
                return new BannerDiagnosticResult(BannerDiagnosticOutcome.NoSelection, bannerSignature.Trim().Length, false);

            string needle = bannerSignature.Trim();
            int idx = htmlBody.IndexOf(needle, StringComparison.Ordinal);
            if (idx >= 0)
                return new BannerDiagnosticResult(BannerDiagnosticOutcome.Found, needle.Length, true);

            var pattern = BuildWhitespaceTolerantPattern(needle);
            if (pattern != null)
            {
                try
                {
                    var m = Regex.Match(htmlBody, pattern, RegexOptions.None, MatchTimeout);
                    if (m.Success)
                        return new BannerDiagnosticResult(BannerDiagnosticOutcome.Found, needle.Length, false);
                }
                catch (RegexMatchTimeoutException) { }
            }

            return new BannerDiagnosticResult(BannerDiagnosticOutcome.NotFound, needle.Length, false);
        }

        // --- internals ------------------------------------------------------

        /// <summary>Tidy a double blank line left where the banner used to sit (cosmetic only).</summary>
        private static string Collapse(string html)
        {
            try
            {
                return Regex.Replace(html, "(\\s*<br[^>]*>\\s*){3,}", "<br><br>",
                                     RegexOptions.IgnoreCase, MatchTimeout);
            }
            catch (RegexMatchTimeoutException) { return html; }
        }

        private static string BuildWhitespaceTolerantPattern(string block)
        {
            // Each maximal non-whitespace run becomes a literal; the gaps between
            // them allow any whitespace, INCLUDING none - the signature may have
            // a newline between two tags where the target body has them adjacent.
            var tokens = Regex.Split(block, "\\s+").Where(t => t.Length > 0).ToList();
            if (tokens.Count == 0) return null;
            return string.Join("\\s*", tokens.Select(Regex.Escape));
        }

        private static int IndexOfOpenTag(string html, int from, string tag)
        {
            var m = Regex.Match(html.Substring(from), "<" + tag + "\\b",
                                RegexOptions.IgnoreCase, MatchTimeout);
            return m.Success ? from + m.Index : -1;
        }

        /// <summary>
        /// Return the first <c>&lt;tag&gt;…&lt;/tag&gt;</c> block at or after
        /// <paramref name="from"/>, matching nested same-name tags so the close
        /// is the right one. Null when no opening tag or no balanced close.
        /// </summary>
        private static string ExtractFirstOfTag(string html, int from, string tag)
        {
            int open = IndexOfOpenTag(html, from, tag);
            if (open < 0) return null;

            string openRe = "<" + tag + "\\b";
            string closeRe = "</" + tag + "\\s*>";
            int depth = 0;
            int pos = open;

            while (pos < html.Length)
            {
                Match next;
                try
                {
                    next = Regex.Match(html.Substring(pos),
                                       "(" + openRe + ")|(" + closeRe + ")",
                                       RegexOptions.IgnoreCase, MatchTimeout);
                }
                catch (RegexMatchTimeoutException) { return null; }

                if (!next.Success) return null;

                int at = pos + next.Index;
                bool isOpen = next.Groups[1].Success;
                if (isOpen)
                {
                    // Ignore a self-closed open tag (<tag ... />) - rare for these.
                    int gt = html.IndexOf('>', at);
                    bool selfClosed = gt > 0 && html[gt - 1] == '/';
                    if (!selfClosed) depth++;
                    pos = (gt > 0 ? gt : at + next.Length) + 1;
                }
                else
                {
                    depth--;
                    int end = at + next.Length;
                    if (depth == 0)
                        return html.Substring(open, end - open);
                    pos = end;
                }
            }

            return null;
        }
    }
}
