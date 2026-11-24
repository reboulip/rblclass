using System;
using System.Formats.Asn1;
using System.Linq;
using System.Numerics;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// The CMS envelope for re-encrypted filing (v2.14.0.0 B1): the shape of the
    /// real Stormshield mail (AES-256-CBC, RSA PKCS#1 v1.5, issuer + serial).
    /// </summary>
    public class CmsEnvelopeBuilderTests
    {
        private sealed class Parsed
        {
            public string ContentType;
            public BigInteger EnvelopedVersion;
            public int RecipientCount;
            public BigInteger RecipientVersion;
            public byte[] Issuer;
            public BigInteger Serial;
            public string KeyTransport;
            public bool KeyTransportHasNullParams;
            public byte[] EncryptedKey;
            public string InnerContentType;
            public string ContentCipher;
            public byte[] Iv;
            public byte[] Ciphertext;
        }

        private static (RSA key, X509Certificate2 cert) NewRecipient(string subject = "CN=RBLclass test recipient")
        {
            var rsa = RSA.Create(2048);
            var req = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var cert = req.CreateSelfSigned(DateTimeOffset.Now.AddDays(-1), DateTimeOffset.Now.AddDays(1));
            return (rsa, cert);
        }

        private static byte[] Build(byte[] content, X509Certificate2 cert)
        {
            byte[] serial = cert.GetSerialNumber();
            Array.Reverse(serial);
            using (var pub = cert.GetRSAPublicKey())
                return CmsEnvelopeBuilder.Build(content, cert.IssuerName.RawData, serial, pub.ExportParameters(false));
        }

        private static Parsed Parse(byte[] der)
        {
            var p = new Parsed();
            var outer = new AsnReader(der, AsnEncodingRules.DER);
            var contentInfo = outer.ReadSequence();
            outer.HasData.Should().BeFalse();
            p.ContentType = contentInfo.ReadObjectIdentifier();
            var explicit0 = contentInfo.ReadSequence(new Asn1Tag(TagClass.ContextSpecific, 0, true));
            var env = explicit0.ReadSequence();
            p.EnvelopedVersion = env.ReadInteger();

            var recipients = env.ReadSetOf();
            var ktri = recipients.ReadSequence();
            p.RecipientCount = 1;
            while (recipients.HasData) { recipients.ReadEncodedValue(); p.RecipientCount++; }
            p.RecipientVersion = ktri.ReadInteger();
            var ias = ktri.ReadSequence();
            p.Issuer = ias.ReadEncodedValue().ToArray();
            p.Serial = ias.ReadInteger();
            var keyAlg = ktri.ReadSequence();
            p.KeyTransport = keyAlg.ReadObjectIdentifier();
            if (keyAlg.HasData) { keyAlg.ReadNull(); p.KeyTransportHasNullParams = true; }
            p.EncryptedKey = ktri.ReadOctetString();

            var eci = env.ReadSequence();
            p.InnerContentType = eci.ReadObjectIdentifier();
            var cipherAlg = eci.ReadSequence();
            p.ContentCipher = cipherAlg.ReadObjectIdentifier();
            p.Iv = cipherAlg.ReadOctetString();
            p.Ciphertext = eci.ReadOctetString(new Asn1Tag(TagClass.ContextSpecific, 0));
            env.HasData.Should().BeFalse("no unprotected attributes");
            return p;
        }

        [Fact]
        public void The_envelope_has_the_stormshield_shape()
        {
            var (_, cert) = NewRecipient();

            var p = Parse(Build(Encoding.UTF8.GetBytes("MIME-Version: 1.0\r\n\r\nbody"), cert));

            p.ContentType.Should().Be("1.2.840.113549.1.7.3");
            p.EnvelopedVersion.Should().Be(0);
            p.RecipientCount.Should().Be(1);
            p.RecipientVersion.Should().Be(0);
            p.KeyTransport.Should().Be("1.2.840.113549.1.1.1", "RSA PKCS#1 v1.5, not OAEP");
            p.KeyTransportHasNullParams.Should().BeTrue();
            p.EncryptedKey.Should().HaveCount(256);
            p.InnerContentType.Should().Be("1.2.840.113549.1.7.1");
            p.ContentCipher.Should().Be("2.16.840.1.101.3.4.1.42");
            p.Iv.Should().HaveCount(16);
        }

        [Fact]
        public void The_recipient_is_the_certificate_issuer_and_serial()
        {
            var (_, cert) = NewRecipient();

            var p = Parse(Build(new byte[] { 1, 2, 3 }, cert));

            p.Issuer.Should().Equal(cert.IssuerName.RawData);
            p.Serial.Should().Be(new BigInteger(cert.GetSerialNumber(), isUnsigned: true, isBigEndian: false));
        }

        [Fact]
        public void The_private_key_decrypts_the_content_back()
        {
            var (rsa, cert) = NewRecipient();
            byte[] content = Encoding.UTF8.GetBytes(new string('é', 5000) + "\r\nfin");

            var p = Parse(Build(content, cert));

            byte[] key = rsa.Decrypt(p.EncryptedKey, RSAEncryptionPadding.Pkcs1);
            key.Should().HaveCount(32);
            using (var aes = Aes.Create())
            {
                aes.Key = key;
                aes.DecryptCbc(p.Ciphertext, p.Iv, PaddingMode.PKCS7).Should().Equal(content);
            }
        }

        [Fact]
        public void Every_envelope_uses_a_fresh_key_and_iv()
        {
            var (_, cert) = NewRecipient();
            byte[] content = Encoding.UTF8.GetBytes("same content");

            var a = Parse(Build(content, cert));
            var b = Parse(Build(content, cert));

            a.Iv.Should().NotEqual(b.Iv);
            a.Ciphertext.Should().NotEqual(b.Ciphertext);
        }

        [Fact]
        public void Long_content_uses_multi_byte_lengths()
        {
            // Long content forces multi-byte DER lengths everywhere.
            var (_, cert) = NewRecipient();
            byte[] content = new byte[70000];
            new Random(42).NextBytes(content);

            var p = Parse(Build(content, cert));

            p.Ciphertext.Length.Should().Be(70000 + 16 - 70000 % 16);
        }

        [Theory]
        [InlineData(new byte[] { 0x00, 0xA5, 0xF4, 0x51 }, "A5F451")]   // the real Stormshield serial shape
        [InlineData(new byte[] { 0xA5, 0xF4 }, "A5F4")]                 // high bit set: sign byte added
        [InlineData(new byte[] { 0x00, 0x00, 0x12 }, "12")]             // redundant zeros dropped
        public void Serials_are_encoded_as_minimal_non_negative_integers(byte[] magnitude, string expectedHex)
        {
            var (_, cert) = NewRecipient();
            byte[] der;
            using (var pub = cert.GetRSAPublicKey())
                der = CmsEnvelopeBuilder.Build(new byte[] { 1 }, cert.IssuerName.RawData, magnitude, pub.ExportParameters(false));

            // The DER reader rejects a non-minimal INTEGER, so parsing proves the encoding.
            var p = Parse(der);
            p.Serial.Should().Be(BigInteger.Parse("0" + expectedHex, System.Globalization.NumberStyles.HexNumber));
        }

        [Fact]
        public void Object_identifiers_with_multi_byte_arcs_round_trip()
        {
            var (_, cert) = NewRecipient();
            var p = Parse(Build(new byte[] { 1 }, cert));
            // 840 and 113549 need several base-128 bytes; 3.4.1.42 under 2.16.840.1.101.
            p.ContentType.Should().Be(CmsEnvelopeBuilder.OidEnvelopedData);
            p.ContentCipher.Should().Be(CmsEnvelopeBuilder.OidAes256Cbc);
        }
    }
}
