using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// Every shipped RBLclass assembly carries the product version from
    /// <c>release.config.json</c>. The MSI replaces a versioned DLL only when
    /// its file version differs from the installed one, so a project left on
    /// an old version keeps its previous build on upgrade. That happened with
    /// the 2.14.0.1 beta: Core and Adapter were still 2.4.0.0, the target kept
    /// the beta1 Adapter and only the add-in shell was updated.
    /// </summary>
    public sealed class VersionParityTests
    {
        private static readonly string RepoRoot = FindRepoRoot();

        private static string FindRepoRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "RBLclass.sln")))
                dir = dir.Parent;

            if (dir == null)
                throw new InvalidOperationException("Could not locate RBLclass.sln above " + AppContext.BaseDirectory);

            return dir.FullName;
        }

        private static string ProductVersion()
        {
            var path = Path.Combine(RepoRoot, ".claude", "skills", "make-release", "release.config.json");
            using (var doc = JsonDocument.Parse(File.ReadAllText(path)))
            {
                return doc.RootElement.GetProperty("AssemblyVersion").GetString();
            }
        }

        private static string AssemblyInfoAttribute(string attribute)
        {
            var path = Path.Combine(RepoRoot, "src", "RBLclass.AddIn", "Properties", "AssemblyInfo.cs");
            var match = Regex.Match(
                File.ReadAllText(path),
                @"^\s*\[assembly:\s*" + attribute + @"\(""([^""]+)""\)\]",
                RegexOptions.Multiline);
            match.Success.Should().BeTrue("AssemblyInfo.cs must declare [assembly: {0}(...)]", attribute);
            return match.Groups[1].Value;
        }

        [Theory]
        [InlineData("AssemblyVersion")]
        [InlineData("AssemblyFileVersion")]
        public void AddIn_assembly_info_carries_the_product_version(string attribute)
        {
            AssemblyInfoAttribute(attribute).Should().Be(ProductVersion(),
                "RBLclass.AddIn's {0} must match release.config.json", attribute);
        }

        [Theory]
        [InlineData("RBLclass.Core", "Version")]
        [InlineData("RBLclass.Core", "AssemblyVersion")]
        [InlineData("RBLclass.Core", "FileVersion")]
        [InlineData("RBLclass.Outlook.Adapter", "Version")]
        [InlineData("RBLclass.Outlook.Adapter", "AssemblyVersion")]
        [InlineData("RBLclass.Outlook.Adapter", "FileVersion")]
        public void Project_carries_the_product_version(string project, string property)
        {
            var path = Path.Combine(RepoRoot, "src", project, project + ".csproj");
            var values = XDocument.Load(path).Root
                .Elements("PropertyGroup")
                .Elements(property)
                .Select(e => e.Value.Trim())
                .ToList();

            values.Should().ContainSingle("{0}.csproj must declare <{1}> once", project, property);
            values[0].Should().Be(ProductVersion(),
                "{0}.csproj <{1}> must match release.config.json", project, property);
        }

        [Fact]
        public void Built_core_assembly_carries_the_product_version()
        {
            var assembly = typeof(ProductInfo).Assembly;

            assembly.GetName().Version.ToString().Should().Be(ProductVersion());
            FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion.Should().Be(ProductVersion());
        }
    }
}
