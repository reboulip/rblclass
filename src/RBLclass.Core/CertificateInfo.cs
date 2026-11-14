using System;

namespace RBLclass.Core
{
    /// <summary>
    /// An encryption certificate as shown in Settings and used by re-encrypted
    /// filing (v2.12.0.0). Public data only: no key material.
    /// </summary>
    public sealed class CertificateInfo
    {
        public CertificateInfo(string thumbprint, string subject, DateTime notAfter, bool hasPrivateKey)
        {
            Thumbprint = thumbprint ?? string.Empty;
            Subject = subject ?? string.Empty;
            NotAfter = notAfter;
            HasPrivateKey = hasPrivateKey;
        }

        /// <summary>Upper-case hex, no spaces.</summary>
        public string Thumbprint { get; }

        /// <summary>A short display name (common name, e-mail).</summary>
        public string Subject { get; }

        /// <summary>Expiry, local time.</summary>
        public DateTime NotAfter { get; }

        /// <summary>
        /// Windows sees the private key, so a copy encrypted to this certificate can
        /// be decrypted with this machine's keys. Automatic requires it.
        /// </summary>
        public bool HasPrivateKey { get; }
    }
}
