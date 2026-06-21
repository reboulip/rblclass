using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class FavoriteFolderServiceTests
    {
        private static FavoriteFolder Fav(string path) =>
            new FavoriteFolder(path, path.TrimEnd('\\').Split('\\').Last());

        /// <summary>A service whose search cache is pre-loaded with the given favourites.</summary>
        private static FavoriteFolderService WithCache(params FavoriteFolder[] cache)
        {
            var repo = Substitute.For<IFavoriteFolderRepository>();
            repo.LoadFavorites().Returns(cache);
            var svc = new FavoriteFolderService(
                Substitute.For<IDirectoryScanner>(), repo, Substitute.For<ISettingsStore>());
            svc.LoadFromCache();
            return svc;
        }

        private static FavoriteFolderService WithRoots(IDirectoryScanner scanner,
                                                       IFavoriteFolderRepository repo,
                                                       string rootsValue)
        {
            var settings = Substitute.For<ISettingsStore>();
            settings.Get(SettingsKeys.AttachmentFavoriteFolders, Arg.Any<string>())
                    .Returns(rootsValue);
            return new FavoriteFolderService(scanner, repo, settings);
        }

        [Fact]
        public void Search_returns_empty_for_a_blank_query()
        {
            var svc = WithCache(Fav(@"C:\Users\Alice\Documents\Contracts"));
            svc.Search("").Results.Should().BeEmpty();
        }

        [Fact]
        public void Search_matches_a_path_segment_case_insensitively()
        {
            var svc = WithCache(Fav(@"C:\Users\Alice\Documents\Contracts"));

            svc.Search("contr").Results.Should().ContainSingle()
               .Which.Folder.FullPath.Should().Be(@"C:\Users\Alice\Documents\Contracts");
        }

        [Fact]
        public void Search_ANDs_across_keywords()
        {
            var svc = WithCache(
                Fav(@"C:\Work\Marketing\2024"),
                Fav(@"C:\Work\Finance\2024"));

            svc.Search("finance 2024").Results.Should().ContainSingle()
               .Which.Folder.FullPath.Should().Be(@"C:\Work\Finance\2024");
        }

        [Fact]
        public void Search_is_accent_insensitive()
        {
            var svc = WithCache(Fav(@"C:\Données\Clients"));
            svc.Search("donnees").Results.Should().ContainSingle();
        }

        [Fact]
        public void Search_on_an_empty_index_returns_empty()
        {
            WithCache().Search("docs").Results.Should().BeEmpty();
        }

        [Fact]
        public void Reindex_indexes_each_root_and_its_subdirectories()
        {
            var scanner = Substitute.For<IDirectoryScanner>();
            var repo = Substitute.For<IFavoriteFolderRepository>();
            scanner.GetDirectories(@"C:\Root", Arg.Any<int>(), Arg.Any<int>())
                   .Returns(new[] { @"C:\Root", @"C:\Root\Sub1", @"C:\Root\Sub2" });

            WithRoots(scanner, repo, @"C:\Root").Reindex();

            repo.Received(1).SaveFavorites(Arg.Is<IEnumerable<FavoriteFolder>>(
                f => f.Select(x => x.Path).SequenceEqual(
                    new[] { @"C:\Root", @"C:\Root\Sub1", @"C:\Root\Sub2" })));
        }

        [Fact]
        public void Reindex_derives_the_display_name_from_the_last_segment()
        {
            var scanner = Substitute.For<IDirectoryScanner>();
            var repo = Substitute.For<IFavoriteFolderRepository>();
            scanner.GetDirectories(@"C:\Root", Arg.Any<int>(), Arg.Any<int>())
                   .Returns(new[] { @"C:\Root\Contracts" });

            WithRoots(scanner, repo, @"C:\Root").Reindex();

            repo.Received(1).SaveFavorites(Arg.Is<IEnumerable<FavoriteFolder>>(
                f => f.Single().DisplayName == "Contracts"));
        }

        [Fact]
        public void Reindex_skips_duplicate_roots()
        {
            var scanner = Substitute.For<IDirectoryScanner>();
            var repo = Substitute.For<IFavoriteFolderRepository>();
            scanner.GetDirectories(@"C:\A", Arg.Any<int>(), Arg.Any<int>())
                   .Returns(new[] { @"C:\A" });

            WithRoots(scanner, repo, @"C:\A;C:\A").Reindex();

            scanner.Received(1).GetDirectories(@"C:\A", Arg.Any<int>(), Arg.Any<int>());
            repo.Received(1).SaveFavorites(Arg.Is<IEnumerable<FavoriteFolder>>(f => f.Count() == 1));
        }

        [Fact]
        public void Reindex_of_a_missing_root_persists_nothing()
        {
            var scanner = Substitute.For<IDirectoryScanner>();
            var repo = Substitute.For<IFavoriteFolderRepository>();
            // The scanner reports a non-existent / inaccessible root as empty.
            scanner.GetDirectories(@"Z:\Nope", Arg.Any<int>(), Arg.Any<int>())
                   .Returns(new string[0]);

            WithRoots(scanner, repo, @"Z:\Nope").Reindex();

            repo.Received(1).SaveFavorites(Arg.Is<IEnumerable<FavoriteFolder>>(f => !f.Any()));
        }

        [Fact]
        public void Reindex_caps_the_total_indexed_directories()
        {
            var scanner = Substitute.For<IDirectoryScanner>();
            var repo = Substitute.For<IFavoriteFolderRepository>();
            // Scanner ignores the budget and returns more than the cap.
            var many = Enumerable.Range(0, FavoriteFolderService.DefaultMaxTotal + 500)
                                 .Select(i => @"C:\Big\d" + i).ToArray();
            scanner.GetDirectories(@"C:\Big", Arg.Any<int>(), Arg.Any<int>()).Returns(many);

            WithRoots(scanner, repo, @"C:\Big").Reindex();

            repo.Received(1).SaveFavorites(Arg.Is<IEnumerable<FavoriteFolder>>(
                f => f.Count() <= FavoriteFolderService.DefaultMaxTotal));
        }

        [Fact]
        public void LoadFromCache_makes_the_persisted_index_searchable_without_a_walk()
        {
            var scanner = Substitute.For<IDirectoryScanner>();
            var repo = Substitute.For<IFavoriteFolderRepository>();
            repo.LoadFavorites().Returns(new[] { Fav(@"C:\Archive\Invoices") });
            var svc = new FavoriteFolderService(scanner, repo, Substitute.For<ISettingsStore>());

            svc.LoadFromCache();

            svc.Search("invoices").Results.Should().ContainSingle();
            scanner.DidNotReceive().GetDirectories(Arg.Any<string>(), Arg.Any<int>(), Arg.Any<int>());
        }
    }
}
