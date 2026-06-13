using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public sealed class UiLanguageResolverTests
    {
        [Theory]
        [InlineData("Auto", "fr", "fr")]
        [InlineData("Auto", "de", "de")]
        [InlineData("Auto", "en", "en")]
        [InlineData("Auto", "es", "en")] // unsupported Outlook language -> English
        [InlineData("Auto", null, "en")] // undetectable Outlook language -> English
        [InlineData("de", "fr", "de")]   // explicit preference wins over Outlook's language
        [InlineData("fr", null, "fr")]   // explicit preference wins even if Outlook is undetectable
        [InlineData("es", "fr", "fr")]   // unrecognised preference falls through to Outlook's language
        [InlineData(null, "de", "de")]   // missing preference falls through to Outlook's language
        [InlineData("", "de", "de")]     // empty preference falls through to Outlook's language
        public void Resolve_picks_the_expected_language(string preferredSetting, string outlookLanguageCode, string expected)
        {
            UiLanguageResolver.Resolve(preferredSetting, outlookLanguageCode).Should().Be(expected);
        }

        [Theory]
        [InlineData("en", true)]
        [InlineData("fr", true)]
        [InlineData("de", true)]
        [InlineData("es", false)]
        [InlineData("Auto", false)]
        [InlineData(null, false)]
        public void IsSupportedLanguage_only_accepts_translated_languages(string languageCode, bool expected)
        {
            UiLanguageResolver.IsSupportedLanguage(languageCode).Should().Be(expected);
        }
    }
}
