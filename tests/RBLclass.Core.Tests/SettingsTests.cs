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
            settings.ClassifySafetyCopy.Should().BeFalse();
            settings.ExternalBannerSignature.Should().BeEmpty();
            settings.StripBannerOnReply.Should().BeFalse();
            settings.StripBannerOnClassify.Should().BeFalse();
            settings.WidenConversation.Should().BeFalse();
            settings.SendExternalWarning.Should().BeTrue();
            settings.InternalDomains.Should().BeEmpty();
            settings.ForgottenAttachmentKeywords.Should().Equal("attach", "enclos", "joint", "PJ");
            settings.SentItemTriageMode.Should().Be(SentItemTriageMode.AskEveryTime);
            settings.ClassifyAfterMoveToInbox.Should().BeTrue();
            settings.AttachmentRemovalMode.Should().Be(AttachmentRemovalMode.Modal);
            settings.AttachmentLabelLocation.Should().Be(AttachmentLabelLocation.Body);
            settings.MinSearchLength.Should().Be(FolderSearchOptions.DefaultMinQueryLength);
            settings.SearchDebounceMs.Should().Be(Settings.DefaultSearchDebounceMs);
            settings.PreferredUiLanguage.Should().Be("Auto");
            settings.AutoExpandResults.Should().BeFalse();
            settings.AutoClassHistoryDays.Should().Be(Settings.DefaultAutoClassHistoryDays);
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
                ClassifySafetyCopy = true,
                ExternalBannerSignature = "<table><tr><td>CAUTION external</td></tr></table>",
                StripBannerOnReply = true,
                StripBannerOnClassify = true,
                WidenConversation = true,
                SendExternalWarning = false,
                InternalDomains = new[] { "example.com", "example.org" },
                ForgottenAttachmentKeywords = new[] { "pièce jointe", "ci-joint" },
                SentItemTriageMode = SentItemTriageMode.Delete,
                ClassifyAfterMoveToInbox = false,
                AttachmentRemovalMode = AttachmentRemovalMode.DeleteSilently,
                AttachmentLabelLocation = AttachmentLabelLocation.InfoBar,
                MinSearchLength = 3,
                SearchDebounceMs = 450,
                AutoExpandResults = true,
                AutoClassHistoryDays = 180,
                PreferredUiLanguage = "fr"
            };

            settings.Save(_store);
            var reloaded = Settings.Load(_store);

            reloaded.OpenInNewWindow.Should().BeTrue();
            reloaded.AllResults.Should().BeTrue();
            reloaded.FolderMatchMode.Should().Be(FolderMatchMode.Substring);
            reloaded.MaxResults.Should().Be(250);
            reloaded.KeepCopy.Should().BeTrue();
            reloaded.RemoveAttachments.Should().BeTrue();
            reloaded.ClassifySafetyCopy.Should().BeTrue();
            reloaded.ExternalBannerSignature.Should().Be("<table><tr><td>CAUTION external</td></tr></table>");
            reloaded.StripBannerOnReply.Should().BeTrue();
            reloaded.StripBannerOnClassify.Should().BeTrue();
            reloaded.WidenConversation.Should().BeTrue();
            reloaded.SendExternalWarning.Should().BeFalse();
            reloaded.InternalDomains.Should().Equal("example.com", "example.org");
            reloaded.ForgottenAttachmentKeywords.Should().Equal("pièce jointe", "ci-joint");
            reloaded.SentItemTriageMode.Should().Be(SentItemTriageMode.Delete);
            reloaded.ClassifyAfterMoveToInbox.Should().BeFalse();
            reloaded.AttachmentRemovalMode.Should().Be(AttachmentRemovalMode.DeleteSilently);
            reloaded.AttachmentLabelLocation.Should().Be(AttachmentLabelLocation.InfoBar);
            reloaded.MinSearchLength.Should().Be(3);
            reloaded.SearchDebounceMs.Should().Be(450);
            reloaded.AutoExpandResults.Should().BeTrue();
            reloaded.AutoClassHistoryDays.Should().Be(180);
            reloaded.PreferredUiLanguage.Should().Be("fr");
        }

        [Theory]
        [InlineData("fr", "fr")]
        [InlineData("de", "de")]
        [InlineData("en", "en")]
        [InlineData("Auto", "Auto")]
        [InlineData("es", "Auto")]          // unsupported -> Auto
        [InlineData("garbage", "Auto")]     // unrecognised -> Auto
        public void Load_normalises_the_preferred_ui_language(string stored, string expected)
        {
            _store.Set(SettingsKeys.PreferredUiLanguage, stored);
            Settings.Load(_store).PreferredUiLanguage.Should().Be(expected);
        }

        [Fact]
        public void Load_defaults_the_preferred_ui_language_to_auto_when_unset()
        {
            Settings.Load(_store).PreferredUiLanguage.Should().Be("Auto");
        }

        [Theory]
        [InlineData("not-a-number", FolderSearchOptions.DefaultMinQueryLength)] // invalid -> default
        [InlineData("0", 1)]    // below range -> clamped to 1
        [InlineData("99", 10)]  // above range -> clamped to MaxMinSearchLength
        [InlineData("4", 4)]    // in range -> kept
        public void Load_clamps_or_defaults_the_min_search_length(string stored, int expected)
        {
            _store.Set(SettingsKeys.MinSearchLength, stored);
            Settings.Load(_store).MinSearchLength.Should().Be(expected);
        }

        [Theory]
        [InlineData("garbage", Settings.DefaultSearchDebounceMs)] // invalid -> default
        [InlineData("-1", 0)]      // below range -> clamped to 0 (no debounce)
        [InlineData("60000", 2000)] // above range -> clamped to MaxSearchDebounceMs
        [InlineData("350", 350)]   // in range -> kept
        public void Load_clamps_or_defaults_the_search_debounce(string stored, int expected)
        {
            _store.Set(SettingsKeys.SearchDebounceMs, stored);
            Settings.Load(_store).SearchDebounceMs.Should().Be(expected);
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
        public void Load_falls_back_to_Modal_when_the_stored_attachment_removal_mode_is_unrecognised()
        {
            _store.Set(SettingsKeys.AttachmentRemovalMode, "Shred");
            Settings.Load(_store).AttachmentRemovalMode.Should().Be(AttachmentRemovalMode.Modal);
        }

        [Fact]
        public void Load_falls_back_to_Body_when_the_stored_attachment_label_location_is_unrecognised()
        {
            _store.Set(SettingsKeys.AttachmentLabelLocation, "Hologram");
            Settings.Load(_store).AttachmentLabelLocation.Should().Be(AttachmentLabelLocation.Body);
        }

        [Fact]
        public void Load_falls_back_to_the_default_substring_when_the_stored_match_mode_is_unrecognised()
        {
            _store.Set(SettingsKeys.FolderMatchMode, "Fuzzy");
            Settings.Load(_store).FolderMatchMode.Should().Be(FolderMatchMode.Substring);
        }

        [Theory]
        [InlineData("not-a-number", Settings.DefaultAutoClassHistoryDays)]
        [InlineData("0", 1)]
        [InlineData("99999", Settings.MaxAutoClassHistoryDays)]
        [InlineData("365", 365)]
        public void Load_clamps_or_defaults_AutoClassHistoryDays(string stored, int expected)
        {
            _store.Set(SettingsKeys.AutoClassHistoryDays, stored);
            Settings.Load(_store).AutoClassHistoryDays.Should().Be(expected);
        }
    }
}
