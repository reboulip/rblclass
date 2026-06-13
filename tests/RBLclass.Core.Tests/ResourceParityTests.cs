using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// Future-proofing guardrail for the i18n sprint: every key in
    /// <c>Strings.resx</c> (English baseline) must also exist, non-empty, in
    /// <c>Strings.fr.resx</c> and <c>Strings.de.resx</c> with the same set of
    /// "{n}" format placeholders. Likewise every ribbon button's
    /// label/screentip/supertip must exist, non-empty, in
    /// <c>Ribbon.fr.xml</c> and <c>Ribbon.de.xml</c>. A new English string
    /// added without translations fails this test at `dotnet test` time.
    /// </summary>
    public sealed class ResourceParityTests
    {
        private static readonly string AddInDir = FindAddInDir();

        private static string FindAddInDir()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "RBLclass.sln")))
                dir = dir.Parent;

            if (dir == null)
                throw new InvalidOperationException("Could not locate RBLclass.sln above " + AppContext.BaseDirectory);

            return Path.Combine(dir.FullName, "src", "RBLclass.AddIn");
        }

        private static readonly Regex PlaceholderRegex = new Regex(@"\{(\d+)\}", RegexOptions.Compiled);

        private static IReadOnlySet<int> Placeholders(string value) =>
            new HashSet<int>(PlaceholderRegex.Matches(value).Select(m => int.Parse(m.Groups[1].Value)));

        private static Dictionary<string, string> LoadResx(string fileName)
        {
            var path = Path.Combine(AddInDir, "Resources", fileName);
            var doc = XDocument.Load(path);
            return doc.Root.Elements("data")
                .ToDictionary(e => e.Attribute("name").Value, e => e.Element("value").Value);
        }

        [Theory]
        [InlineData("Strings.fr.resx")]
        [InlineData("Strings.de.resx")]
        public void Translation_has_same_keys_as_English_baseline(string translatedFile)
        {
            var en = LoadResx("Strings.resx");
            var translated = LoadResx(translatedFile);

            translated.Keys.Should().BeEquivalentTo(en.Keys,
                "{0} must define exactly the same resource keys as Strings.resx", translatedFile);
        }

        [Theory]
        [InlineData("Strings.fr.resx")]
        [InlineData("Strings.de.resx")]
        public void Translation_values_are_non_empty(string translatedFile)
        {
            var translated = LoadResx(translatedFile);

            foreach (var kvp in translated)
                kvp.Value.Trim().Should().NotBeEmpty($"{translatedFile} key '{kvp.Key}' must not be empty");
        }

        [Theory]
        [InlineData("Strings.fr.resx")]
        [InlineData("Strings.de.resx")]
        public void Translation_placeholders_match_English_baseline(string translatedFile)
        {
            var en = LoadResx("Strings.resx");
            var translated = LoadResx(translatedFile);

            foreach (var kvp in en)
            {
                if (!translated.TryGetValue(kvp.Key, out var translatedValue)) continue; // reported by the key-parity test

                Placeholders(translatedValue).Should().BeEquivalentTo(Placeholders(kvp.Value),
                    $"key '{kvp.Key}' must use the same {{n}} placeholders in {translatedFile} as in Strings.resx");
            }
        }

        private const string RibbonNamespace = "http://schemas.microsoft.com/office/2009/07/customui";

        private static Dictionary<string, Dictionary<string, string>> LoadRibbonButtons(string fileName)
        {
            XNamespace ns = RibbonNamespace;
            var path = Path.Combine(AddInDir, fileName);
            var doc = XDocument.Load(path);

            return doc.Descendants(ns + "button")
                .ToDictionary(
                    b => b.Attribute("id").Value,
                    b => new[] { "label", "screentip", "supertip" }
                        .ToDictionary(attr => attr, attr => b.Attribute(attr)?.Value ?? string.Empty));
        }

        [Theory]
        [InlineData("Ribbon.fr.xml")]
        [InlineData("Ribbon.de.xml")]
        public void Ribbon_translation_has_same_buttons_and_attributes_as_English_baseline(string translatedFile)
        {
            var en = LoadRibbonButtons("Ribbon.xml");
            var translated = LoadRibbonButtons(translatedFile);

            translated.Keys.Should().BeEquivalentTo(en.Keys,
                "{0} must define the same set of ribbon buttons as Ribbon.xml", translatedFile);

            foreach (var id in en.Keys)
            {
                foreach (var attr in en[id].Keys)
                {
                    translated[id][attr].Trim().Should().NotBeEmpty(
                        $"{translatedFile} button '{id}' attribute '{attr}' must not be empty");
                }
            }
        }
    }
}
