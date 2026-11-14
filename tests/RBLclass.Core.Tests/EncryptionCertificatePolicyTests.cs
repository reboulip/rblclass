using System;
using System.Linq;
using System.Text;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class EncryptionCertificatePolicyTests
    {
        private static readonly DateTime Now = new DateTime(2026, 11, 14, 12, 0, 0);

        private static CertificateCandidate Cert(string thumb = "AA11", string email = "me@corp.fr",
                                                 bool privateKey = true, bool rsa = true, bool? keyEnc = true,
                                                 string[] ekus = null, int expiresInDays = 365, int startedDaysAgo = 30,
                                                 bool noEku = false) =>
            new CertificateCandidate(thumb, "CN=Me", Now.AddDays(-startedDaysAgo), Now.AddDays(expiresInDays),
                                     privateKey, rsa, keyEnc,
                                     noEku ? null : (ekus ?? new[] { EncryptionCertificatePolicy.SecureEmailEku }),
                                     email == null ? null : new[] { email });

        [Fact]
        public void A_valid_secure_email_rsa_certificate_is_a_candidate()
        {
            EncryptionCertificatePolicy.IsCandidate(Cert(), Now).Should().BeTrue();
        }

        [Fact]
        public void Missing_key_usage_and_eku_extensions_are_allowed()
        {
            EncryptionCertificatePolicy.IsCandidate(Cert(keyEnc: null, noEku: true), Now).Should().BeTrue();
        }

        [Fact]
        public void Any_purpose_eku_is_allowed()
        {
            EncryptionCertificatePolicy.IsCandidate(Cert(ekus: new[] { "2.5.29.37.0" }), Now).Should().BeTrue();
        }

        [Theory]
        [InlineData("expired")]
        [InlineData("notyet")]
        [InlineData("notrsa")]
        [InlineData("nokeyenc")]
        [InlineData("clientauth")]
        public void Unsuitable_certificates_are_not_candidates(string flaw)
        {
            var cert =
                flaw == "expired" ? Cert(expiresInDays: -1) :
                flaw == "notyet" ? Cert(startedDaysAgo: -1) :
                flaw == "notrsa" ? Cert(rsa: false) :
                flaw == "nokeyenc" ? Cert(keyEnc: false) :
                Cert(ekus: new[] { "1.3.6.1.5.5.7.3.2" });

            EncryptionCertificatePolicy.IsCandidate(cert, Now).Should().BeFalse();
        }

        [Fact]
        public void Automatic_picks_the_matching_address_with_the_latest_expiry()
        {
            var older = Cert("01", expiresInDays: 100);
            var newer = Cert("02", expiresInDays: 500);
            var other = Cert("03", email: "someone@else.fr", expiresInDays: 900);

            EncryptionCertificatePolicy.SelectAutomatic(new[] { older, newer, other }, " ME@corp.fr ", Now)
                .Thumbprint.Should().Be("02");
        }

        [Fact]
        public void Automatic_requires_a_visible_private_key()
        {
            var noKey = Cert("01", privateKey: false, expiresInDays: 900);
            var withKey = Cert("02", expiresInDays: 100);

            EncryptionCertificatePolicy.SelectAutomatic(new[] { noKey, withKey }, "me@corp.fr", Now)
                .Thumbprint.Should().Be("02");
            EncryptionCertificatePolicy.SelectAutomatic(new[] { noKey }, "me@corp.fr", Now).Should().BeNull();
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("nobody@corp.fr")]
        public void Automatic_finds_nothing_without_a_matching_account(string smtp)
        {
            EncryptionCertificatePolicy.SelectAutomatic(new[] { Cert() }, smtp, Now).Should().BeNull();
        }

        [Fact]
        public void An_explicit_choice_may_pick_a_certificate_without_a_visible_private_key()
        {
            var noKey = Cert("ab cd", privateKey: false);

            EncryptionCertificatePolicy.ResolveExplicit(new[] { noKey }, "‎ABCD", Now)
                .Thumbprint.Should().Be("ABCD");
        }

        [Fact]
        public void An_explicit_choice_that_is_no_longer_valid_resolves_to_nothing()
        {
            var expired = Cert("ABCD", expiresInDays: -1);
            var valid = Cert("EF01");

            EncryptionCertificatePolicy.ResolveExplicit(new[] { expired, valid }, "ABCD", Now).Should().BeNull();
        }

        [Fact]
        public void Candidates_are_filtered_and_ordered_by_expiry()
        {
            var list = EncryptionCertificatePolicy.Candidates(
                new[] { Cert("01", expiresInDays: 10), Cert("02", rsa: false), Cert("03", expiresInDays: 50) }, Now);

            list.Select(c => c.Thumbprint).Should().Equal("03", "01");
        }

        [Theory]
        [InlineData("ab cd ef", "ABCDEF")]
        [InlineData("‎AB:CD", "ABCD")]
        [InlineData(null, "")]
        public void Thumbprints_are_normalised(string raw, string expected)
        {
            EncryptionCertificatePolicy.NormalizeThumbprint(raw).Should().Be(expected);
        }

        [Fact]
        public void Rfc822_names_are_read_from_the_san_der()
        {
            // SEQUENCE { [2] dNSName "a.b", [1] rfc822Name "me@corp.fr", [1] "x@y.z" }
            byte[] dns = { 0x82, 3, (byte)'a', (byte)'.', (byte)'b' };
            byte[] mail1 = new byte[] { 0x81, 10 }.Concat(Encoding.ASCII.GetBytes("me@corp.fr")).ToArray();
            byte[] mail2 = new byte[] { 0x81, 5 }.Concat(Encoding.ASCII.GetBytes("x@y.z")).ToArray();
            byte[] body = dns.Concat(mail1).Concat(mail2).ToArray();
            byte[] der = new byte[] { 0x30, (byte)body.Length }.Concat(body).ToArray();

            EncryptionCertificatePolicy.ParseRfc822Names(der).Should().Equal("me@corp.fr", "x@y.z");
        }

        [Fact]
        public void Malformed_san_der_yields_no_exception()
        {
            EncryptionCertificatePolicy.ParseRfc822Names(null).Should().BeEmpty();
            EncryptionCertificatePolicy.ParseRfc822Names(new byte[] { 0x30, 0x7F, 0x81, 0x50 }).Should().BeEmpty();
            EncryptionCertificatePolicy.ParseRfc822Names(new byte[] { 0x04, 0x00 }).Should().BeEmpty();
        }
    }
}
