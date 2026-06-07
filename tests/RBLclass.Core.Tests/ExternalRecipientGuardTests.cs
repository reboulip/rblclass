using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class ExternalRecipientGuardTests
    {
        private static RecipientAddress Exchange(string address) => new RecipientAddress("Colleague", address, isExchangeResolved: true);
        private static RecipientAddress Smtp(string address) => new RecipientAddress("Contact", address, isExchangeResolved: false);

        [Fact]
        public void Exchange_resolved_recipients_are_never_external()
        {
            var sut = new ExternalRecipientGuard();

            var external = sut.FindExternal(new[] { Exchange("alice@contoso.com") }, internalDomains: new string[0]);

            external.Should().BeEmpty();
        }

        [Fact]
        public void Smtp_recipients_on_an_allow_listed_domain_are_internal()
        {
            var sut = new ExternalRecipientGuard();

            var external = sut.FindExternal(
                new[] { Smtp("bob@contoso.com") },
                internalDomains: new[] { "contoso.com" });

            external.Should().BeEmpty();
        }

        [Fact]
        public void Domain_matching_is_case_insensitive()
        {
            var sut = new ExternalRecipientGuard();

            var external = sut.FindExternal(
                new[] { Smtp("bob@Contoso.COM") },
                internalDomains: new[] { "CONTOSO.com" });

            external.Should().BeEmpty();
        }

        [Fact]
        public void Smtp_recipients_off_the_allowlist_are_external()
        {
            var sut = new ExternalRecipientGuard();
            var outsider = Smtp("eve@example.org");

            var external = sut.FindExternal(
                new[] { outsider },
                internalDomains: new[] { "contoso.com" });

            external.Should().BeEquivalentTo(new[] { outsider });
        }

        [Fact]
        public void Empty_allowlist_treats_every_non_exchange_recipient_as_external()
        {
            var sut = new ExternalRecipientGuard();
            var outsider = Smtp("eve@example.org");

            var external = sut.FindExternal(new[] { outsider }, internalDomains: new string[0]);

            external.Should().BeEquivalentTo(new[] { outsider });
        }

        [Fact]
        public void Reports_only_the_external_subset_of_a_mixed_recipient_list()
        {
            var sut = new ExternalRecipientGuard();
            var colleague = Exchange("alice@contoso.com");
            var partner = Smtp("bob@partner.com");
            var outsider = Smtp("eve@example.org");

            var external = sut.FindExternal(
                new[] { colleague, partner, outsider },
                internalDomains: new[] { "partner.com" });

            external.Should().BeEquivalentTo(new[] { outsider });
        }

        [Fact]
        public void Recipient_with_no_address_domain_is_external_unless_exchange_resolved()
        {
            var sut = new ExternalRecipientGuard();
            var malformed = Smtp("not-an-address");

            var external = sut.FindExternal(new[] { malformed }, internalDomains: new[] { "contoso.com" });

            external.Should().BeEquivalentTo(new[] { malformed });
        }
    }
}
