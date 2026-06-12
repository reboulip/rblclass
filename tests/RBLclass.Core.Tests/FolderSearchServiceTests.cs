using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class FolderSearchServiceTests
    {
        // A small fixed corpus under one store "s1":
        //   Archive / Projects               (e1, non-leaf)
        //   Archive / Projects / ProjetJuridique (e2, leaf)
        //   Archive / Projects / Marketing   (e3, leaf)
        //   Archive / Administration         (e4, leaf)
        //   Archive / Finance                (e5, non-leaf)
        //   Archive / Finance / Factures 2024 (e6, leaf)
        //   Archive / Règlement              (e7, leaf, accented)
        private static readonly FolderNode[] Corpus =
        {
            new FolderNode("s1", "e1", null, "Projects", "Archive / Projects", isLeaf: false),
            new FolderNode("s1", "e2", "e1", "ProjetJuridique", "Archive / Projects / ProjetJuridique", isLeaf: true),
            new FolderNode("s1", "e3", "e1", "Marketing", "Archive / Projects / Marketing", isLeaf: true),
            new FolderNode("s1", "e4", null, "Administration", "Archive / Administration", isLeaf: true),
            new FolderNode("s1", "e5", null, "Finance", "Archive / Finance", isLeaf: false),
            new FolderNode("s1", "e6", "e5", "Factures 2024", "Archive / Finance / Factures 2024", isLeaf: true),
            new FolderNode("s1", "e7", null, "Règlement", "Archive / Règlement", isLeaf: true),
        };

        private static IFolderSearch SearchOver(params FolderNode[] nodes)
        {
            var tree = Substitute.For<IFolderTree>();
            tree.GetAll().Returns(nodes);
            return new FolderSearchService(tree);
        }

        private static IFolderSearch DefaultSearch() => SearchOver(Corpus);

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("\t")]
        public void Blank_query_returns_empty(string query)
        {
            var outcome = DefaultSearch().Search(query);

            outcome.Results.Should().BeEmpty();
            outcome.TotalMatchCount.Should().Be(0);
            outcome.LimitExceeded.Should().BeFalse();
        }

        [Fact]
        public void No_match_returns_empty()
        {
            DefaultSearch().Search("zzz").Results.Should().BeEmpty();
        }

        [Fact]
        public void WordPrefix_matched_non_leaf_collapses_to_topmost()
        {
            // "proj" matches Projects and (via the "Projects" path word) both of
            // its children; only the topmost (Projects) is returned, collapsed.
            var outcome = DefaultSearch().Search("proj");

            outcome.Results.Should().ContainSingle();
            var r = outcome.Results[0];
            r.Folder.EntryId.Should().Be("e1");
            r.IsCollapsed.Should().BeTrue();
            r.DisplayPath.Should().Be("Archive / Projects" + FolderSearchResult.CollapsedSuffix);
        }

        [Fact]
        public void AllResults_expands_every_matching_folder()
        {
            var outcome = DefaultSearch().Search("proj",
                new FolderSearchOptions(allResults: true));

            outcome.Results.Select(r => r.Folder.EntryId)
                   .Should().Equal("e1", "e3", "e2"); // alphabetical by full path
            outcome.Results.All(r => !r.IsCollapsed).Should().BeTrue();
        }

        [Fact]
        public void Leaf_match_is_not_collapsed()
        {
            var outcome = DefaultSearch().Search("administration");

            outcome.Results.Should().ContainSingle();
            outcome.Results[0].Folder.EntryId.Should().Be("e4");
            outcome.Results[0].IsCollapsed.Should().BeFalse();
            outcome.Results[0].DisplayPath.Should().Be("Archive / Administration");
        }

        [Fact]
        public void And_across_keywords_with_one_keyword_from_an_ancestor_segment()
        {
            // "archive" is an ancestor segment, "factures" is in the leaf only.
            var outcome = DefaultSearch().Search("archive factures");

            outcome.Results.Should().ContainSingle()
                   .Which.Folder.EntryId.Should().Be("e6");
        }

        [Fact]
        public void And_requires_all_keywords()
        {
            // "finance" matches, "marketing" does not share a path with it.
            DefaultSearch().Search("finance marketing").Results.Should().BeEmpty();
        }

        [Fact]
        public void Default_substring_matches_mid_word_while_word_prefix_opt_in_does_not()
        {
            // "juri" is in the middle of the single word "ProjetJuridique".
            // The default (substring) matches it...
            DefaultSearch().Search("juri").Results.Should().ContainSingle()
                           .Which.Folder.EntryId.Should().Be("e2");

            // ...while the opt-in word-prefix mode does not.
            DefaultSearch().Search("juri",
                new FolderSearchOptions(matchMode: FolderMatchMode.WordPrefix))
                .Results.Should().BeEmpty();
        }

        [Fact]
        public void Default_match_mode_is_substring()
        {
            FolderSearchOptions.Default.MatchMode.Should().Be(FolderMatchMode.Substring);
            new FolderSearchOptions().MatchMode.Should().Be(FolderMatchMode.Substring);
        }

        [Fact]
        public void Default_search_matches_a_keyword_inside_a_word()
        {
            // Pilot example: "security" should find "Cybersecurity" out of the box.
            var search = SearchOver(
                new FolderNode("s1", "c1", null, "Cybersecurity", "Archive / Cybersecurity", isLeaf: true));

            search.Search("security").Results.Should().ContainSingle()
                  .Which.Folder.EntryId.Should().Be("c1");
        }

        [Theory]
        [InlineData(FolderMatchMode.Substring)]
        [InlineData(FolderMatchMode.WordPrefix)]
        public void Keyword_with_special_character_matches_in_both_modes(FolderMatchMode mode)
        {
            // Pilot bug: searching "R&D" did not find the folder named "R&D"
            // (word-prefix split paths into letter/digit words, so "r&d" could
            // never prefix one).
            var search = SearchOver(
                new FolderNode("s1", "rd1", null, "R&D", "Archive / R&D", isLeaf: true),
                new FolderNode("s1", "rd2", null, "Research", "Archive / Research", isLeaf: true));

            search.Search("R&D", new FolderSearchOptions(matchMode: mode))
                  .Results.Should().ContainSingle()
                  .Which.Folder.EntryId.Should().Be("rd1");
        }

        [Fact]
        public void Special_character_keyword_combines_with_plain_word_prefix_keywords()
        {
            var search = SearchOver(
                new FolderNode("s1", "rd1", null, "R&D", "Archive / Projects / R&D", isLeaf: true),
                new FolderNode("s1", "rd2", null, "R&D", "Personal / R&D", isLeaf: true));

            // "proj" stays a word prefix (AND semantics across keywords).
            var outcome = search.Search("proj r&d",
                new FolderSearchOptions(matchMode: FolderMatchMode.WordPrefix));

            outcome.Results.Should().ContainSingle()
                   .Which.Folder.EntryId.Should().Be("rd1");
        }

        [Fact]
        public void Plain_keywords_keep_word_prefix_strictness()
        {
            // The special-character fallback must not loosen ordinary keywords:
            // "juri" (mid-word) still misses in word-prefix mode.
            DefaultSearch().Search("juri",
                new FolderSearchOptions(matchMode: FolderMatchMode.WordPrefix))
                .Results.Should().BeEmpty();
        }

        [Fact]
        public void Matching_is_accent_and_case_insensitive()
        {
            DefaultSearch().Search("reglement").Results.Should().ContainSingle()
                           .Which.Folder.EntryId.Should().Be("e7");
            DefaultSearch().Search("RÈGLEMENT").Results.Should().ContainSingle()
                           .Which.Folder.EntryId.Should().Be("e7");
        }

        [Fact]
        public void Results_are_sorted_alphabetically_by_full_path()
        {
            var outcome = DefaultSearch().Search("archive",
                new FolderSearchOptions(allResults: true));

            var paths = outcome.Results.Select(r => r.Folder.FullPath).ToList();
            paths.Should().BeInAscendingOrder();
            paths.First().Should().Be("Archive / Administration");
        }

        [Fact]
        public void Over_cap_truncates_and_flags_limit_exceeded()
        {
            var outcome = DefaultSearch().Search("archive",
                new FolderSearchOptions(allResults: true, maxResults: 3));

            outcome.TotalMatchCount.Should().Be(7); // every folder path has "Archive"
            outcome.LimitExceeded.Should().BeTrue();
            outcome.Results.Should().HaveCount(3);
            // still the first 3 alphabetically
            outcome.Results.Select(r => r.Folder.FullPath)
                   .Should().Equal("Archive / Administration",
                                   "Archive / Finance",
                                   "Archive / Finance / Factures 2024");
        }

        [Fact]
        public void Topmost_collapse_spans_only_the_matching_subtree()
        {
            // "finance" matches Finance + its child; collapses to Finance.
            var outcome = DefaultSearch().Search("finance");

            outcome.Results.Should().ContainSingle();
            outcome.Results[0].Folder.EntryId.Should().Be("e5");
            outcome.Results[0].IsCollapsed.Should().BeTrue();
        }
    }
}
