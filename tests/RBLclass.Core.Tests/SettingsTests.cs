using System;
using System.IO;
using FluentAssertions;
using RBLclass.Core;
using RBLclass.Core.Persistence;
using Xunit;

namespace RBLclass.Core.Tests
{
    public sealed class SettingsTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly SqliteSettingsStore _store;

        public SettingsTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(),
                "rblclass-settings-" + Guid.NewGuid().ToString("N") + ".db");
            _store = new SqliteSettingsStore("Data Source=" + _dbPath);
            _store.EnsureSchema();
        }

        public void Dispose()
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
            foreach (var p in new[] { _dbPath, _dbPath + "-wal", _dbPath + "-shm" })
            {
                try { if (File.Exists(p)) File.Delete(p); } catch { }
            }
        }

        [Fact]
        public void Load_on_an_empty_store_returns_the_documented_defaults()
        {
            var settings = Settings.Load(_store);

            settings.OpenInNewWindow.Should().BeFalse();
            settings.AllResults.Should().BeFalse();
            settings.FolderMatchMode.Should().Be(FolderMatchMode.Substring);
            settings.MaxResults.Should().Be(FolderSearchOptions.DefaultMaxResults);
            settings.KeepCopy.Should().BeFalse();
            settings.RemoveAttachments.Should().BeFalse();
            settings.WidenConversation.Should().BeFalse();
            settings.SendExternalWarning.Should().BeTrue();
            settings.InternalDomains.Should().BeEmpty();
            settings.ForgottenAttachmentKeywords.Should().Equal("attach", "enclos", "joint", "PJ");
            settings.SentItemTriageMode.Should().Be(SentItemTriageMode.AskEveryTime);
        }

        [Fact]
        public void Save_then_Load_round_trips_every_field()
        {
            var settings = new Settings
            {
                OpenInNewWindow = true,
                AllResults = true,
                FolderMatchMode = FolderMatchMode.Substring,
                MaxResults = 250,
                KeepCopy = true,
                RemoveAttachments = true,
                WidenConversation = true,
                SendExternalWarning = false,
                InternalDomains = new[] { "example.com", "example.org" },
                ForgottenAttachmentKeywords = new[] { "pièce jointe", "ci-joint" },
                SentItemTriageMode = SentItemTriageMode.Delete
            };

            settings.Save(_store);
            var reloaded = Settings.Load(_store);

            reloaded.OpenInNewWindow.Should().BeTrue();
            reloaded.AllResults.Should().BeTrue();
            reloaded.FolderMatchMode.Should().Be(FolderMatchMode.Substring);
            reloaded.MaxResults.Should().Be(250);
            reloaded.KeepCopy.Should().BeTrue();
            reloaded.RemoveAttachments.Should().BeTrue();
            reloaded.WidenConversation.Should().BeTrue();
            reloaded.SendExternalWarning.Should().BeFalse();
            reloaded.InternalDomains.Should().Equal("example.com", "example.org");
            reloaded.ForgottenAttachmentKeywords.Should().Equal("pièce jointe", "ci-joint");
            reloaded.SentItemTriageMode.Should().Be(SentItemTriageMode.Delete);
        }

        [Theory]
        [InlineData(true, SentItemTriageMode.AskEveryTime)]
        [InlineData(false, SentItemTriageMode.Leave)]
        public void Load_migrates_the_legacy_triage_prompt_flag_when_no_mode_is_stored(
            bool legacyPromptOn, SentItemTriageMode expected)
        {
            _store.SetBool(SettingsKeys.SentItemTriagePrompt, legacyPromptOn);

            Settings.Load(_store).SentItemTriageMode.Should().Be(expected);
        }

        [Fact]
        public void Load_prefers_a_stored_triage_mode_over_the_legacy_flag()
        {
            _store.SetBool(SettingsKeys.SentItemTriagePrompt, false); // legacy says "leave"
            _store.Set(SettingsKeys.SentItemTriageMode, SentItemTriageMode.MoveToInbox.ToString());

            Settings.Load(_store).SentItemTriageMode.Should().Be(SentItemTriageMode.MoveToInbox);
        }

        [Fact]
        public void Save_trims_blank_entries_out_of_lists()
        {
            var settings = Settings.Load(_store);
            settings.InternalDomains = new[] { " example.com ", "", "  ", "example.org" };

            settings.Save(_store);

            _store.Get(SettingsKeys.InternalDomains).Should().Be("example.com;example.org");
            Settings.Load(_store).InternalDomains.Should().Equal("example.com", "example.org");
        }

        [Fact]
        public void Load_falls_back_to_the_default_max_results_when_the_stored_value_is_invalid()
        {
            _store.Set(SettingsKeys.MaxResults, "not-a-number");
            Settings.Load(_store).MaxResults.Should().Be(FolderSearchOptions.DefaultMaxResults);

            _store.Set(SettingsKeys.MaxResults, "0");
            Settings.Load(_store).MaxResults.Should().Be(FolderSearchOptions.DefaultMaxResults);

            _store.Set(SettingsKeys.MaxResults, "-5");
            Settings.Load(_store).MaxResults.Should().Be(FolderSearchOptions.DefaultMaxResults);
        }

        [Fact]
        public void Load_falls_back_to_the_default_substring_when_the_stored_match_mode_is_unrecognised()
        {
            _store.Set(SettingsKeys.FolderMatchMode, "Fuzzy");
            Settings.Load(_store).FolderMatchMode.Should().Be(FolderMatchMode.Substring);
        }
    }
}
