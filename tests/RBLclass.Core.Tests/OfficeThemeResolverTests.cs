using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class OfficeThemeResolverTests
    {
        [Theory]
        [InlineData(4)] // Black
        [InlineData(3)] // Dark Gray
        public void Explicit_dark_office_codes_resolve_dark_regardless_of_windows(int office)
        {
            OfficeThemeResolver.Resolve(office, windowsAppsUseLightTheme: true)
                .Should().Be(ThemeMode.Dark);
            OfficeThemeResolver.Resolve(office, windowsAppsUseLightTheme: null)
                .Should().Be(ThemeMode.Dark);
        }

        [Theory]
        [InlineData(5)] // White
        [InlineData(0)] // Colorful (classic builds)
        [InlineData(7)] // Colorful (Current channel builds, e.g. the dev machine)
        public void Explicit_light_office_codes_resolve_light_regardless_of_windows(int office)
        {
            OfficeThemeResolver.Resolve(office, windowsAppsUseLightTheme: false)
                .Should().Be(ThemeMode.Light);
            OfficeThemeResolver.Resolve(office, windowsAppsUseLightTheme: null)
                .Should().Be(ThemeMode.Light);
        }

        [Theory]
        [InlineData(6)]    // "Use system setting" on older builds
        [InlineData(99)]   // any future/unknown code
        [InlineData(null)] // absent
        public void Unknown_or_system_codes_follow_windows_dark(int? office)
        {
            OfficeThemeResolver.Resolve(office, windowsAppsUseLightTheme: false)
                .Should().Be(ThemeMode.Dark);
        }

        [Theory]
        [InlineData(6)]
        [InlineData(99)]
        [InlineData(null)]
        public void Unknown_or_system_codes_follow_windows_light(int? office)
        {
            OfficeThemeResolver.Resolve(office, windowsAppsUseLightTheme: true)
                .Should().Be(ThemeMode.Light);
        }

        [Fact]
        public void Defaults_to_light_when_everything_is_unknown()
        {
            OfficeThemeResolver.Resolve(officeUiTheme: null, windowsAppsUseLightTheme: null)
                .Should().Be(ThemeMode.Light);
        }
    }
}
