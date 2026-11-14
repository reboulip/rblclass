using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RBLclass.Core
{
    /// <summary>
    /// Writes a decrypted <see cref="MailContent"/> as one MIME entity, the content
    /// that re-encrypted filing wraps in S/MIME (v2.12.0.0 A2). No RFC 822
    /// envelope headers: those stay outside the encryption.
    /// <para>
    /// Structure: the body (<c>text/plain</c>, <c>text/html</c>, or
    /// <c>multipart/alternative</c> when both exist), wrapped in
    /// <c>multipart/related</c> with the inline parts that carry a content id
    /// (HTML only), wrapped in <c>multipart/mixed</c> with every other part. A
    /// level with a single node collapses to that node.
    /// </para>
    /// <para>
    /// Everything is base64 (bodies too): netstandard2.0 has no public
    /// quoted-printable encoder, base64 meets the 76/998 line rules by
    /// construction, and its alphabet has no '-', so a <c>--boundary</c> line can
    /// never occur inside an encoded body: only header text needs checking
    /// against the boundaries. Parts are encoded in chunks into buffers kept off
    /// the large-object heap.
    /// </para>
    /// <para>
    /// Header values that come from the sender's mail (file names, content ids,
    /// content types) are sanitised, so a CR/LF in them cannot inject headers.
    /// </para>
    /// </summary>
    public sealed class MimeEntityBuilder
    {
        private const int MaxBoundaryAttempts = 10;
        private const int Base64LineChars = 76;
        private const int Base64LineBytes = 57;           // 57 bytes -> 76 base64 chars
        private const int ChunkLines = 500;               // 28,500 bytes in, 38,000 chars out
        private const int SimpleNameMaxLength = 60;
        private const int Rfc2231SegmentMax = 60;
        private const int EncodedWordMaxBytes = 45;       // 45 bytes -> 60 base64 chars -> 72-char word
        private const string OctetStream = "application/octet-stream";

        private static readonly Encoding Utf8 = new UTF8Encoding(false);

        private static readonly Regex ContentTypePattern = new Regex(
            @"^[a-z0-9!#$&^_.+\-]+/[a-z0-9!#$&^_.+\-]+$",
            RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(250));

        private static readonly Regex MetaCharset = new Regex(
            @"(?<pre>charset\s*=\s*)(?<q>[""']?)(?<v>[A-Za-z0-9_.:\-]+)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(250));

        private readonly Func<string> _boundaryFactory;

        /// <param name="boundaryFactory">Test seam; defaults to <c>=_RBL_</c> + a new GUID.</param>
        public MimeEntityBuilder(Func<string> boundaryFactory = null)
        {
            _boundaryFactory = boundaryFactory ?? (() => "=_RBL_" + Guid.NewGuid().ToString("N"));
        }

        public byte[] Build(MailContent content)
        {
            var root = BuildTree(content);
            using (var stream = new MemoryStream(EstimateSize(root)))
            {
                Write(root, stream);
                return stream.ToArray();
            }
        }

        public void WriteTo(MailContent content, Stream output)
        {
            if (output == null) throw new ArgumentNullException(nameof(output));
            Write(BuildTree(content), output);
        }

        // ---- tree ------------------------------------------------------------

        private sealed class Node
        {
            public readonly List<string> Headers = new List<string>();  // continuation lines start with TAB
            public byte[] Data;                                        // leaf
            public List<Node> Children;                                // multipart
            public string Subtype;
            public string RelatedType;
            public string Boundary;
            public bool IsMultipart => Children != null;
            public string ContentType;                                 // for related type=
        }

        private Node BuildTree(MailContent content)
        {
            if (content == null) throw new ArgumentNullException(nameof(content));

            string html = string.IsNullOrEmpty(content.HtmlBody) ? null : RewriteMetaCharset(content.HtmlBody);
            string plain = string.IsNullOrEmpty(content.PlainBody) ? null : content.PlainBody;

            Node body;
            if (html != null && plain != null)
                body = Multipart("alternative", new List<Node> { TextLeaf("plain", plain), TextLeaf("html", html) });
            else if (html != null)
                body = TextLeaf("html", html);
            else
                body = TextLeaf("plain", plain ?? string.Empty);

            var related = new List<Node>();
            var mixed = new List<Node>();
            foreach (var part in content.Parts)
            {
                string cid = SanitizeHeaderText(part.ContentId);
                bool toRelated = html != null && part.IsInline && !string.IsNullOrEmpty(cid);
                (toRelated ? related : mixed).Add(PartLeaf(part, cid));
            }

            if (related.Count > 0)
            {
                var children = new List<Node> { body };
                children.AddRange(related);
                var node = Multipart("related", children);
                node.RelatedType = body.ContentType;
                body = node;
            }
            if (mixed.Count > 0)
            {
                var children = new List<Node> { body };
                children.AddRange(mixed);
                body = Multipart("mixed", children);
            }

            AssignBoundaries(body);
            FinishMultipartHeaders(body);
            body.Headers.Insert(0, "MIME-Version: 1.0");
            return body;
        }

        private static Node Multipart(string subtype, List<Node> children) =>
            new Node { Subtype = subtype, Children = children, ContentType = "multipart/" + subtype };

        private static Node TextLeaf(string subtype, string text)
        {
            var node = new Node { ContentType = "text/" + subtype, Data = Utf8.GetBytes(NormalizeLineEnds(text)) };
            node.Headers.Add("Content-Type: text/" + subtype + "; charset=\"utf-8\"");
            node.Headers.Add("Content-Transfer-Encoding: base64");
            return node;
        }

        private static Node PartLeaf(MailPart part, string cid)
        {
            string type = SanitizeContentType(part.ContentType);
            string name = SanitizeHeaderText(part.FileName);
            var node = new Node { ContentType = type, Data = part.Data };

            if (string.IsNullOrEmpty(name))
            {
                node.Headers.Add("Content-Type: " + type);
            }
            else
            {
                node.Headers.Add("Content-Type: " + type + ";");
                node.Headers.AddRange(NameParameter(name));
            }
            node.Headers.Add("Content-Transfer-Encoding: base64");

            string disposition = part.IsInline ? "inline" : "attachment";
            if (string.IsNullOrEmpty(name))
            {
                node.Headers.Add("Content-Disposition: " + disposition);
            }
            else
            {
                node.Headers.Add("Content-Disposition: " + disposition + ";");
                node.Headers.AddRange(FilenameParameter(name));
            }

            if (!string.IsNullOrEmpty(cid))
                node.Headers.Add("Content-ID: <" + cid + ">");
            return node;
        }

        // ---- parameters (RFC 2231 / RFC 2047) ----------------------------------

        private static bool IsSimpleName(string name)
        {
            if (name.Length > SimpleNameMaxLength) return false;
            foreach (char c in name)
                if (c < 0x20 || c > 0x7E || c == '"' || c == '\\') return false;
            return true;
        }

        private static IEnumerable<string> NameParameter(string name)
        {
            if (IsSimpleName(name)) return new[] { "\tname=\"" + name + "\"" };

            var words = EncodedWords(name).ToList();
            var lines = new List<string>();
            for (int i = 0; i < words.Count; i++)
            {
                string line = "\t" + (i == 0 ? "name=\"" : string.Empty) + words[i];
                if (i == words.Count - 1) line += "\"";
                lines.Add(line);
            }
            return lines;
        }

        private static IEnumerable<string> FilenameParameter(string name)
        {
            if (IsSimpleName(name)) return new[] { "\tfilename=\"" + name + "\"" };

            var segments = Rfc2231Segments(name);
            if (segments.Count == 1) return new[] { "\tfilename*=" + segments[0] };

            var lines = new List<string>();
            for (int i = 0; i < segments.Count; i++)
                lines.Add("\tfilename*" + i + "*=" + segments[i] + (i < segments.Count - 1 ? ";" : string.Empty));
            return lines;
        }

        /// <summary>RFC 2231 extended value split into segments that never break a %XX triplet.</summary>
        private static List<string> Rfc2231Segments(string value)
        {
            var tokens = new List<string>();
            foreach (byte b in Utf8.GetBytes(value))
            {
                char c = (char)b;
                bool attrChar = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')
                                || "!#$&+-.^_`|~".IndexOf(c) >= 0;
                tokens.Add(attrChar ? c.ToString() : "%" + b.ToString("X2"));
            }

            var segments = new List<string>();
            var current = new StringBuilder("utf-8''");   // charset prefix: first segment only
            int tokensInCurrent = 0;
            foreach (string token in tokens)
            {
                if (tokensInCurrent > 0 && current.Length + token.Length > Rfc2231SegmentMax)
                {
                    segments.Add(current.ToString());
                    current.Clear();
                    tokensInCurrent = 0;
                }
                current.Append(token);
                tokensInCurrent++;
            }
            segments.Add(current.ToString());
            return segments;
        }

        /// <summary>RFC 2047 B-encoded words of at most 45 UTF-8 bytes, never splitting a character.</summary>
        private static IEnumerable<string> EncodedWords(string value)
        {
            var chunk = new StringBuilder();
            int chunkBytes = 0;
            for (int i = 0; i < value.Length; i++)
            {
                string unit = char.IsHighSurrogate(value[i]) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1])
                    ? value.Substring(i++, 2)
                    : value[i].ToString();
                int bytes = Utf8.GetByteCount(unit);
                if (chunkBytes + bytes > EncodedWordMaxBytes && chunk.Length > 0)
                {
                    yield return Word(chunk.ToString());
                    chunk.Clear();
                    chunkBytes = 0;
                }
                chunk.Append(unit);
                chunkBytes += bytes;
            }
            if (chunk.Length > 0) yield return Word(chunk.ToString());
        }

        private static string Word(string text) => "=?utf-8?B?" + Convert.ToBase64String(Utf8.GetBytes(text)) + "?=";

        // ---- boundaries --------------------------------------------------------

        private void AssignBoundaries(Node root)
        {
            var multiparts = new List<Node>();
            var headerText = new StringBuilder();
            Collect(root, multiparts, headerText);
            string headers = headerText.ToString();

            var chosen = new List<string>();
            foreach (var node in multiparts)
            {
                string boundary = null;
                for (int attempt = 0; attempt < MaxBoundaryAttempts && boundary == null; attempt++)
                {
                    string candidate = _boundaryFactory();
                    if (IsUsableBoundary(candidate, headers, chosen)) boundary = candidate;
                }
                if (boundary == null)
                    throw new InvalidOperationException("No usable MIME boundary after " + MaxBoundaryAttempts + " attempts.");
                chosen.Add(boundary);
                node.Boundary = boundary;
            }
        }

        private static bool IsUsableBoundary(string candidate, string headers, List<string> chosen)
        {
            if (string.IsNullOrEmpty(candidate) || candidate.Length > 70) return false;
            foreach (char c in candidate)
                if (c < 0x20 || c > 0x7E || c == '"' || c == '\\') return false;
            if (headers.IndexOf(candidate, StringComparison.Ordinal) >= 0) return false;
            // A boundary that is a prefix of another would match the other's delimiter lines.
            return chosen.All(b => b.IndexOf(candidate, StringComparison.Ordinal) < 0
                                   && candidate.IndexOf(b, StringComparison.Ordinal) < 0);
        }

        private static void Collect(Node node, List<Node> multiparts, StringBuilder headerText)
        {
            foreach (string header in node.Headers) headerText.Append(header).Append('\n');
            if (!node.IsMultipart) return;
            multiparts.Add(node);
            foreach (var child in node.Children) Collect(child, multiparts, headerText);
        }

        private static void FinishMultipartHeaders(Node node)
        {
            if (!node.IsMultipart) return;
            node.Headers.Add("Content-Type: multipart/" + node.Subtype + ";");
            if (node.RelatedType != null)
            {
                node.Headers.Add("\tboundary=\"" + node.Boundary + "\";");
                node.Headers.Add("\ttype=\"" + node.RelatedType + "\"");
            }
            else
            {
                node.Headers.Add("\tboundary=\"" + node.Boundary + "\"");
            }
            foreach (var child in node.Children) FinishMultipartHeaders(child);
        }

        // ---- writing -----------------------------------------------------------

        private static void Write(Node root, Stream output)
        {
            WriteNode(root, output);
            WriteAscii(output, "\r\n");
        }

        /// <summary>Headers, blank line, body - without a trailing CRLF (the caller adds the delimiter).</summary>
        private static void WriteNode(Node node, Stream output)
        {
            foreach (string header in node.Headers) WriteAscii(output, header + "\r\n");
            WriteAscii(output, "\r\n");

            if (!node.IsMultipart)
            {
                WriteBase64(node.Data, output);
                return;
            }

            foreach (var child in node.Children)
            {
                WriteAscii(output, "--" + node.Boundary + "\r\n");
                WriteNode(child, output);
                WriteAscii(output, "\r\n");
            }
            WriteAscii(output, "--" + node.Boundary + "--");
        }

        private static void WriteBase64(byte[] data, Stream output)
        {
            const int chunkBytes = Base64LineBytes * ChunkLines;
            var chars = new char[ChunkLines * Base64LineChars];
            var bytes = new byte[ChunkLines * (Base64LineChars + 2)];
            bool first = true;

            for (int offset = 0; offset < data.Length; offset += chunkBytes)
            {
                int length = Math.Min(chunkBytes, data.Length - offset);
                int charCount = Convert.ToBase64CharArray(data, offset, length, chars, 0);
                int written = 0;
                for (int lineStart = 0; lineStart < charCount; lineStart += Base64LineChars)
                {
                    if (!first)
                    {
                        bytes[written++] = (byte)'\r';
                        bytes[written++] = (byte)'\n';
                    }
                    first = false;
                    int lineLength = Math.Min(Base64LineChars, charCount - lineStart);
                    for (int i = 0; i < lineLength; i++) bytes[written++] = (byte)chars[lineStart + i];
                }
                output.Write(bytes, 0, written);
            }
        }

        private static void WriteAscii(Stream output, string text)
        {
            var bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        private static int EstimateSize(Node node)
        {
            long size = 2;
            foreach (string header in node.Headers) size += header.Length + 2;
            if (node.IsMultipart)
            {
                foreach (var child in node.Children) size += node.Boundary.Length + 6 + EstimateSize(child);
                size += node.Boundary.Length + 4;
            }
            else
            {
                long b64 = 4L * ((node.Data.Length + 2) / 3);
                long lines = (b64 + Base64LineChars - 1) / Base64LineChars;
                size += b64 + 2 * Math.Max(lines - 1, 0);
            }
            return (int)Math.Min(int.MaxValue - 1024, size + 64);
        }

        // ---- sanitising --------------------------------------------------------

        private static string NormalizeLineEnds(string text) =>
            text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");

        /// <summary>Drops control characters (CR/LF included) and angle brackets from header text.</summary>
        private static string SanitizeHeaderText(string value)
        {
            if (value == null) return null;
            var sb = new StringBuilder(value.Length);
            foreach (char c in value)
                if (c >= 0x20 && c != 0x7F && c != '<' && c != '>') sb.Append(c);
            return sb.ToString().Trim();
        }

        private static string SanitizeContentType(string contentType)
        {
            if (string.IsNullOrWhiteSpace(contentType)) return OctetStream;
            string type = contentType;
            int semicolon = type.IndexOf(';');
            if (semicolon >= 0) type = type.Substring(0, semicolon);
            type = type.Trim().ToLowerInvariant();
            try { return ContentTypePattern.IsMatch(type) ? type : OctetStream; }
            catch (RegexMatchTimeoutException) { return OctetStream; }
        }

        /// <summary>
        /// Rewrites the charset declared in &lt;meta&gt; tags to utf-8, since the body
        /// is written as UTF-8 (a stale windows-1252 declaration garbles accents).
        /// Tags are found by a linear scan; the regex only runs on one tag's text.
        /// </summary>
        internal static string RewriteMetaCharset(string html)
        {
            var sb = new StringBuilder(html.Length);
            int pos = 0;
            while (true)
            {
                int start = html.IndexOf("<meta", pos, StringComparison.OrdinalIgnoreCase);
                if (start < 0) break;
                int end = html.IndexOf('>', start);
                if (end < 0) break;

                sb.Append(html, pos, start - pos);
                string tag = html.Substring(start, end - start + 1);
                try { tag = MetaCharset.Replace(tag, m => m.Groups["pre"].Value + m.Groups["q"].Value + "utf-8"); }
                catch (RegexMatchTimeoutException) { /* leave the tag as it was */ }
                sb.Append(tag);
                pos = end + 1;
            }
            sb.Append(html, pos, html.Length - pos);
            return sb.ToString();
        }
    }
}
