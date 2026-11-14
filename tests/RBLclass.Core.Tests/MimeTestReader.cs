using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// A minimal MIME parser for <see cref="MimeEntityBuilder"/> tests: unfolds
    /// headers, splits multiparts on their boundary, decodes base64 bodies, RFC
    /// 2231 file names and RFC 2047 names. Test-only; not a general parser.
    /// </summary>
    internal sealed class MimeNode
    {
        public Dictionary<string, string> Headers { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string MediaType { get; set; }
        public Dictionary<string, string> TypeParams { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string Disposition { get; set; }
        public Dictionary<string, string> DispositionParams { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<MimeNode> Children { get; } = new List<MimeNode>();
        public byte[] Body { get; set; }

        public string Text => Encoding.UTF8.GetString(Body);
        public string ContentId => Headers.TryGetValue("Content-ID", out var v) ? v : null;
        public string FileName => MimeTestReader.DecodeFileName(DispositionParams);
        public string Name => TypeParams.TryGetValue("name", out var v) ? MimeTestReader.DecodeEncodedWords(v) : null;
    }

    internal static class MimeTestReader
    {
        public static MimeNode Parse(byte[] entity) => ParseText(Encoding.ASCII.GetString(entity));

        private static MimeNode ParseText(string text)
        {
            int split = text.IndexOf("\r\n\r\n", StringComparison.Ordinal);
            if (split < 0) throw new FormatException("No header/body separator.");
            string headerBlock = text.Substring(0, split);
            string body = text.Substring(split + 4);

            var node = new MimeNode();
            string unfolded = Regex.Replace(headerBlock, "\r\n[ \t]", " ");
            foreach (string line in unfolded.Split(new[] { "\r\n" }, StringSplitOptions.None))
            {
                int colon = line.IndexOf(':');
                if (colon <= 0) throw new FormatException("Bad header line: " + line);
                node.Headers[line.Substring(0, colon).Trim()] = line.Substring(colon + 1).Trim();
            }

            node.TypeParams = ParseParams(node.Headers["Content-Type"], out string mediaType);
            node.MediaType = mediaType;
            if (node.Headers.TryGetValue("Content-Disposition", out string disposition))
            {
                node.DispositionParams = ParseParams(disposition, out string kind);
                node.Disposition = kind;
            }

            if (mediaType.StartsWith("multipart/", StringComparison.Ordinal))
            {
                string boundary = node.TypeParams["boundary"];
                string[] segments = ("\r\n" + body).Split(new[] { "\r\n--" + boundary }, StringSplitOptions.None);
                for (int i = 1; i < segments.Length; i++)
                {
                    if (segments[i].StartsWith("--", StringComparison.Ordinal)) break;
                    node.Children.Add(ParseText(segments[i].Substring(2)));
                }
            }
            else
            {
                node.Body = Convert.FromBase64String(body.Replace("\r\n", string.Empty));
            }
            return node;
        }

        private static Dictionary<string, string> ParseParams(string value, out string head)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var pieces = new List<string>();
            var current = new StringBuilder();
            bool quoted = false;
            foreach (char c in value)
            {
                if (c == '"') quoted = !quoted;
                if (c == ';' && !quoted) { pieces.Add(current.ToString()); current.Clear(); continue; }
                current.Append(c);
            }
            pieces.Add(current.ToString());

            head = pieces[0].Trim().ToLowerInvariant();
            foreach (string piece in pieces.Skip(1))
            {
                int eq = piece.IndexOf('=');
                if (eq < 0) continue;
                string v = piece.Substring(eq + 1).Trim();
                if (v.Length >= 2 && v[0] == '"' && v[v.Length - 1] == '"') v = v.Substring(1, v.Length - 2);
                result[piece.Substring(0, eq).Trim()] = v;
            }
            return result;
        }

        public static string DecodeFileName(Dictionary<string, string> p)
        {
            if (p.TryGetValue("filename", out string simple)) return simple;
            string extended;
            if (!p.TryGetValue("filename*", out extended))
            {
                var parts = new StringBuilder();
                for (int i = 0; p.TryGetValue("filename*" + i + "*", out string seg); i++) parts.Append(seg);
                if (parts.Length == 0) return null;
                extended = parts.ToString();
            }
            int quote = extended.IndexOf("''", StringComparison.Ordinal);
            string encoded = extended.Substring(quote + 2);
            var bytes = new List<byte>();
            for (int i = 0; i < encoded.Length; i++)
            {
                if (encoded[i] == '%')
                {
                    bytes.Add(byte.Parse(encoded.Substring(i + 1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                    i += 2;
                }
                else bytes.Add((byte)encoded[i]);
            }
            return Encoding.UTF8.GetString(bytes.ToArray());
        }

        public static string DecodeEncodedWords(string value)
        {
            var words = Regex.Matches(value, @"=\?utf-8\?B\?(?<b>[A-Za-z0-9+/=]*)\?=");
            if (words.Count == 0) return value;
            var sb = new StringBuilder();
            foreach (Match m in words) sb.Append(Encoding.UTF8.GetString(Convert.FromBase64String(m.Groups["b"].Value)));
            return sb.ToString();
        }
    }
}
