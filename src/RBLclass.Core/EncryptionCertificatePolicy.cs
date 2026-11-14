using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RBLclass.Core
{
    /// <summary>
    /// What the adapter reads from one certificate of the user's store, for
    /// <see cref="EncryptionCertificatePolicy"/> (v2.12.0.0 A5). Outlook- and
    /// crypto-free so the selection rules are unit-tested.
    /// </summary>
    public sealed class CertificateCandidate
    {
        /// <param name="hasKeyEncipherment">Null when the certificate has no Key Usage extension.</param>
        /// <param name="enhancedKeyUsages">Null when it has no EKU extension.</param>
        public CertificateCandidate(string thumbprint, string subject, DateTime notBefore, DateTime notAfter,
                                    bool hasPrivateKey, bool isRsa, bool? hasKeyEncipherment,
                                    IReadOnlyList<string> enhancedKeyUsages, IReadOnlyList<string> emailAddresses)
        {
            Thumbprint = EncryptionCertificatePolicy.NormalizeThumbprint(thumbprint);
            Subject = subject ?? string.Empty;
            NotBefore = notBefore;
            NotAfter = notAfter;
            HasPrivateKey = hasPrivateKey;
            IsRsa = isRsa;
            HasKeyEncipherment = hasKeyEncipherment;
            EnhancedKeyUsages = enhancedKeyUsages;
            EmailAddresses = emailAddresses ?? new string[0];
        }

        public string Thumbprint { get; }
        public string Subject { get; }
        public DateTime NotBefore { get; }
        public DateTime NotAfter { get; }
        public bool HasPrivateKey { get; }
        public bool IsRsa { get; }
        public bool? HasKeyEncipherment { get; }
        public IReadOnlyList<string> EnhancedKeyUsages { get; }
        public IReadOnlyList<string> EmailAddresses { get; }

        public CertificateInfo ToInfo() => new CertificateInfo(Thumbprint, Subject, NotAfter, HasPrivateKey);
    }

    /// <summary>
    /// Which certificates re-encrypted filing may encrypt to (v2.12.0.0 A5).
    /// A candidate is valid now (validity window only: no chain building, which
    /// could fetch revocation lists over the network), has an RSA key, allows key
    /// encipherment (or has no Key Usage extension) and allows secure e-mail (or
    /// has no EKU extension). Automatic also requires an e-mail address matching
    /// the account and a private key Windows can see, so the copy is guaranteed
    /// decryptable here; an explicit choice may pick any candidate.
    /// </summary>
    public static class EncryptionCertificatePolicy
    {
        public const string SecureEmailEku = "1.3.6.1.5.5.7.3.4";
        private const string AnyEku = "2.5.29.37.0";

        public static bool IsCandidate(CertificateCandidate certificate, DateTime now)
        {
            if (certificate == null || !certificate.IsRsa) return false;
            if (now < certificate.NotBefore || now > certificate.NotAfter) return false;
            if (certificate.HasKeyEncipherment == false) return false;
            var ekus = certificate.EnhancedKeyUsages;
            return ekus == null || ekus.Contains(SecureEmailEku) || ekus.Contains(AnyEku);
        }

        public static IReadOnlyList<CertificateCandidate> Candidates(IEnumerable<CertificateCandidate> all, DateTime now) =>
            (all ?? Enumerable.Empty<CertificateCandidate>())
                .Where(c => IsCandidate(c, now))
                .OrderByDescending(c => c.NotAfter)
                .ToList();

        /// <summary>The automatic choice: matching address, private key visible, latest expiry; or null.</summary>
        public static CertificateCandidate SelectAutomatic(IEnumerable<CertificateCandidate> all, string accountSmtp, DateTime now)
        {
            if (string.IsNullOrWhiteSpace(accountSmtp)) return null;
            string smtp = accountSmtp.Trim();
            return Candidates(all, now).FirstOrDefault(c =>
                c.HasPrivateKey
                && c.EmailAddresses.Any(a => string.Equals(a?.Trim(), smtp, StringComparison.OrdinalIgnoreCase)));
        }

        /// <summary>The explicitly chosen certificate when it is still a candidate; never a substitute.</summary>
        public static CertificateCandidate ResolveExplicit(IEnumerable<CertificateCandidate> all, string thumbprint, DateTime now)
        {
            string wanted = NormalizeThumbprint(thumbprint);
            if (wanted.Length == 0) return null;
            return Candidates(all, now).FirstOrDefault(c => c.Thumbprint == wanted);
        }

        /// <summary>Upper-case hex only (drops spaces and the invisible marks a copy from certmgr adds).</summary>
        public static string NormalizeThumbprint(string thumbprint)
        {
            if (string.IsNullOrEmpty(thumbprint)) return string.Empty;
            var sb = new StringBuilder(thumbprint.Length);
            foreach (char c in thumbprint)
                if ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))
                    sb.Append(char.ToUpperInvariant(c));
            return sb.ToString();
        }

        /// <summary>
        /// The rfc822Name entries of a Subject Alternative Name extension value
        /// (DER: SEQUENCE of GeneralName, rfc822Name = [1] IA5String, tag 0x81).
        /// Parsed directly because the framework's formatted output is localised.
        /// Malformed input yields what was read so far.
        /// </summary>
        public static IReadOnlyList<string> ParseRfc822Names(byte[] sanDer)
        {
            var names = new List<string>();
            if (sanDer == null || sanDer.Length < 2 || sanDer[0] != 0x30) return names;

            int pos = 1;
            if (!TryReadLength(sanDer, ref pos, out int seqLength)) return names;
            int end = Math.Min(sanDer.Length, pos + seqLength);
            while (pos < end)
            {
                byte tag = sanDer[pos++];
                if (!TryReadLength(sanDer, ref pos, out int length) || pos + length > end) break;
                if (tag == 0x81) names.Add(Encoding.ASCII.GetString(sanDer, pos, length));
                pos += length;
            }
            return names;
        }

        private static bool TryReadLength(byte[] der, ref int pos, out int length)
        {
            length = 0;
            if (pos >= der.Length) return false;
            byte first = der[pos++];
            if (first < 0x80) { length = first; return true; }
            int count = first & 0x7F;
            if (count == 0 || count > 3 || pos + count > der.Length) return false;
            for (int i = 0; i < count; i++) length = (length << 8) | der[pos++];
            return true;
        }
    }
}
