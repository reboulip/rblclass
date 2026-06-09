using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class FolderExclusionPolicyTests
    {
        private static FolderExclusionPolicy Policy(FolderExclusionOptions options)
            => new FolderExclusionPolicy(options);

        [Fact]
        public void Default_indexes_mailbox_and_pst_but_excludes_public_folders()
        {
            var policy = Policy(FolderExclusionOptions.Default);

            // PST archive - indexed.
            policy.IsStoreExcluded(
                    new StoreInfo("s1", "My Archive", isDataFileStore: true))
                  .Should().BeFalse();
            // Primary mailbox (Exchange/IMAP/OST, not a data file) - indexed.
            policy.IsStoreExcluded(
                    new StoreInfo("s2", "me@corp", isDataFileStore: false))
                  .Should().BeFalse();
            // Public-folder store - excluded.
            policy.IsStoreExcluded(
                    new StoreInfo("s3", "Public Folders", isDataFileStore: false,
                                  isPublicFolderStore: true))
                  .Should().BeTrue();
        }

        [Fact]
        public void ExcludePublicFolderStores_disabled_keeps_public_folders()
        {
            var policy = Policy(new FolderExclusionOptions(excludePublicFolderStores: false));

            policy.IsStoreExcluded(
                    new StoreInfo("s3", "Public Folders", isDataFileStore: false,
                                  isPublicFolderStore: true))
                  .Should().BeFalse();
        }

        [Theory]
        [InlineData("Public Folders - me@corp", true)]
        [InlineData("public folders", true)]   // case-insensitive
        [InlineData("G-FRA archive", true)]
        [InlineData("Project Archive 2024", false)]
        public void Excludes_stores_by_configured_name_substring(string name, bool excluded)
        {
            var policy = Policy(new FolderExclusionOptions(
                excludedStoreNameSubstrings: new[] { "Public Folders", "G-FRA" }));

            policy.IsStoreExcluded(new StoreInfo("s", name, isDataFileStore: true))
                  .Should().Be(excluded);
        }

        [Fact]
        public void Default_excludes_deleted_items_folder()
        {
            var policy = Policy(FolderExclusionOptions.Default);

            policy.IsFolderExcluded("Deleted Items", WellKnownFolderKind.DeletedItems)
                  .Should().BeTrue();
            policy.IsFolderExcluded("Inbox", WellKnownFolderKind.Normal)
                  .Should().BeFalse();
        }

        [Fact]
        public void ExcludeDeletedItems_disabled_keeps_deleted_items()
        {
            var policy = Policy(new FolderExclusionOptions(excludeDeletedItems: false));

            policy.IsFolderExcluded("Deleted Items", WellKnownFolderKind.DeletedItems)
                  .Should().BeFalse();
        }
    }
}
