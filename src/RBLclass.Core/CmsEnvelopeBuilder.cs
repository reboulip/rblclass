using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;

namespace RBLclass.Core
{
    /// <summary>
    /// Builds a CMS <c>EnvelopedData</c> (RFC 5652) in the shape Stormshield uses
    /// (v2.14.0.0 B1, GH #62): content AES-256-CBC, one
    /// <c>KeyTransRecipientInfo</c> addressed by issuer and serial number, key
    /// transport <b>RSA PKCS#1 v1.5</b> (<c>rsaEncryption</c>). The .NET Framework
    /// 4.8 <c>EnvelopedCms</c> emits RSA-OAEP and offers no way to choose the
    /// padding, and a real target crashed on the OAEP copies.
    /// </summary>
    /// <remarks>Pure managed code, DER with definite lengths; no Outlook, no certificate store.</remarks>
    public static class CmsEnvelopeBuilder
    {
        public const string OidEnvelopedData = "1.2.840.113549.1.7.3";
        public const string OidData = "1.2.840.113549.1.7.1";
        public const string OidRsaEncryption = "1.2.840.113549.1.1.1";
        public const string OidAes256Cbc = "2.16.840.1.101.3.4.1.42";

        /// <summary>
        /// The DER <c>ContentInfo</c> wrapping the encrypted <paramref name="content"/>.
        /// </summary>
        /// <param name="content">The plaintext (a MIME entity).</param>
        /// <param name="issuerNameDer">The recipient certificate's issuer Name, DER.</param>
        /// <param name="serialNumberBigEndian">The recipient certificate's serial number, big-endian.</param>
        /// <param name="recipientPublicKey">The recipient's RSA public key.</param>
        public static byte[] Build(byte[] content, byte[] issuerNameDer, byte[] serialNumberBigEndian,
                                   RSAParameters recipientPublicKey)
        {
            if (content == null) throw new ArgumentNullException(nameof(content));
            if (issuerNameDer == null || issuerNameDer.Length == 0) throw new ArgumentException("Issuer name required.", nameof(issuerNameDer));
            if (serialNumberBigEndian == null || serialNumberBigEndian.Length == 0) throw new ArgumentException("Serial number required.", nameof(serialNumberBigEndian));

            byte[] key = new byte[32];
            byte[] iv = new byte[16];
            byte[] encryptedKey;
            byte[] ciphertext;
            try
            {
                using (var rng = RandomNumberGenerator.Create())
                {
                    rng.GetBytes(key);
                    rng.GetBytes(iv);
                }

                using (var aes = Aes.Create())
                {
                    aes.KeySize = 256;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;
                    using (var encryptor = aes.CreateEncryptor(key, iv))
                        ciphertext = encryptor.TransformFinalBlock(content, 0, content.Length);
                }

                using (var rsa = RSA.Create())
                {
                    rsa.ImportParameters(recipientPublicKey);
                    encryptedKey = rsa.Encrypt(key, RSAEncryptionPadding.Pkcs1);
                }

                byte[] recipientInfo = Sequence(
                    Integer(new byte[] { 0 }),                                  // version 0
                    Sequence(issuerNameDer, Integer(serialNumberBigEndian)),    // IssuerAndSerialNumber
                    Sequence(ObjectId(OidRsaEncryption), Null()),               // keyEncryptionAlgorithm
                    OctetString(encryptedKey));

                byte[] encryptedContentInfo = Sequence(
                    ObjectId(OidData),
                    Sequence(ObjectId(OidAes256Cbc), OctetString(iv)),
                    Tagged(0x80, ciphertext));                                  // [0] IMPLICIT OCTET STRING

                byte[] envelopedData = Sequence(
                    Integer(new byte[] { 0 }),                                  // version 0
                    Set(recipientInfo),
                    encryptedContentInfo);

                return Sequence(ObjectId(OidEnvelopedData), Tagged(0xA0, envelopedData));
            }
            finally
            {
                Array.Clear(key, 0, key.Length);
                Array.Clear(iv, 0, iv.Length);
            }
        }

        // ---- DER -------------------------------------------------------------------

        private static byte[] Sequence(params byte[][] parts) => Tagged(0x30, Concat(parts));

        private static byte[] Set(params byte[][] parts) => Tagged(0x31, Concat(parts));

        private static byte[] OctetString(byte[] value) => Tagged(0x04, value);

        private static byte[] Null() => new byte[] { 0x05, 0x00 };

        /// <summary>A non-negative INTEGER from big-endian magnitude bytes, minimally encoded.</summary>
        private static byte[] Integer(byte[] bigEndian)
        {
            int start = 0;
            while (start < bigEndian.Length - 1 && bigEndian[start] == 0 && (bigEndian[start + 1] & 0x80) == 0)
                start++;
            bool pad = (bigEndian[start] & 0x80) != 0;
            var value = new byte[bigEndian.Length - start + (pad ? 1 : 0)];
            Buffer.BlockCopy(bigEndian, start, value, pad ? 1 : 0, bigEndian.Length - start);
            return Tagged(0x02, value);
        }

        private static byte[] ObjectId(string dotted)
        {
            string[] arcs = dotted.Split('.');
            var body = new List<byte> { (byte)(int.Parse(arcs[0]) * 40 + int.Parse(arcs[1])) };
            for (int i = 2; i < arcs.Length; i++)
            {
                long arc = long.Parse(arcs[i]);
                var chunk = new Stack<byte>();
                chunk.Push((byte)(arc & 0x7F));
                arc >>= 7;
                while (arc > 0)
                {
                    chunk.Push((byte)(0x80 | (arc & 0x7F)));
                    arc >>= 7;
                }
                body.AddRange(chunk);
            }
            return Tagged(0x06, body.ToArray());
        }

        private static byte[] Tagged(byte tag, byte[] content)
        {
            using (var ms = new MemoryStream(content.Length + 6))
            {
                ms.WriteByte(tag);
                WriteLength(ms, content.Length);
                ms.Write(content, 0, content.Length);
                return ms.ToArray();
            }
        }

        private static void WriteLength(Stream s, int length)
        {
            if (length < 0x80)
            {
                s.WriteByte((byte)length);
                return;
            }
            var bytes = new List<byte>();
            for (int v = length; v > 0; v >>= 8) bytes.Insert(0, (byte)(v & 0xFF));
            s.WriteByte((byte)(0x80 | bytes.Count));
            foreach (byte b in bytes) s.WriteByte(b);
        }

        private static byte[] Concat(byte[][] parts)
        {
            int total = 0;
            foreach (var p in parts) total += p.Length;
            var result = new byte[total];
            int offset = 0;
            foreach (var p in parts)
            {
                Buffer.BlockCopy(p, 0, result, offset, p.Length);
                offset += p.Length;
            }
            return result;
        }
    }
}
