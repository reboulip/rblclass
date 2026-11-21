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
        public void Query_shorter_than_the_minimum_length_returns_empty()
        {
            // Default minimum is 2: a single keystroke must not search yet.
            DefaultSearch().Search("a").Results.Should().BeEmpty();
            DefaultSearch().Search(" a ").Results.Should().BeEmpty(); // trimmed length counts

            // A custom minimum applies too.
            DefaultSearch().Search("proj",
                new FolderSearchOptions(minQueryLength: 5)).Results.Should().BeEmpty();
        }

        [Fact]
        public void Query_at_the_minimum_length_searches()
        {
            var search = SearchOver(
                new FolderNode("s1", "x1", null, "IT", "Archive / IT", isLeaf: true));

            search.Search("it").Results.Should().ContainSingle()
                  .Which.Folder.EntryId.Should().Be("x1");
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

        [Fact]
        public void Repeated_search_returns_identical_results()
        {
            // The normalized-path cache (v2.7.1.0 E3) must not change results
            // between successive searches over the same tree generation.
            var search = DefaultSearch();

            var first = search.Search("proj");
            var second = search.Search("proj");

            second.TotalMatchCount.Should().Be(first.TotalMatchCount);
            second.Results.Select(r => r.Folder.EntryId)
                  .Should().Equal(first.Results.Select(r => r.Folder.EntryId));
        }

        [Fact]
        public void Search_reflects_the_new_tree_after_a_cache_version_change()
        {
            // First generation: the full corpus. Second: a single new accented
            // folder, with a bumped CacheVersion. The cache must rebuild and the
            // accented match must still normalize correctly (v2.7.1.0 E3).
            var newNode = new FolderNode("s1", "e9", null, "Sécurité",
                                         "Archive / Sécurité", isLeaf: true);
            var tree = Substitute.For<IFolderTree>();
            tree.GetAll().Returns(Corpus, new[] { newNode });
            tree.CacheVersion.Returns(0, 1);
            var search = new FolderSearchService(tree);

            search.Search("projects").Results.Should().NotBeEmpty(); // gen 0

            var outcome = search.Search("securite"); // gen 1
            outcome.Results.Should().ContainSingle()
                   .Which.Folder.EntryId.Should().Be("e9");
        }

        // ---- MaxDepth (v2.13.0.0 B1) -------------------------------------

        // A store root node (null parent, path == name), a top-level folder, its
        // child and grandchild, mirroring the adapter's output.
        private static readonly FolderNode[] DepthTree =
        {
            new FolderNode("s1", "root", null, "Archive", "Archive", isLeaf: false),
            new FolderNode("s1", "a", null, "Alpha", "Archive / Alpha", isLeaf: false),
            new FolderNode("s1", "b", "a", "Beta", "Archive / Alpha / Beta", isLeaf: false),
            new FolderNode("s1", "c", "b", "Gamma", "Archive / Alpha / Beta / Gamma", isLeaf: true),
        };

        private static string[] Paths(FolderSearchOutcome o) =>
            o.Results.Select(r => r.Folder.EntryId).ToArray();

        [Fact]
        public void MaxDepth_zero_means_unlimited()
        {
            var outcome = SearchOver(DepthTree).Search("archive",
                new FolderSearchOptions(allResults: true, maxDepth: 0));

            Paths(outcome).Should().BeEquivalentTo("root", "a", "b", "c");
        }

        [Fact]
        public void Negative_MaxDepth_clamps_to_unlimited()
        {
            new FolderSearchOptions(maxDepth: -3).MaxDepth.Should().Be(0);
        }

        [Fact]
        public void Store_root_is_level_zero_and_its_children_level_one()
        {
            var outcome = SearchOver(DepthTree).Search("archive",
                new FolderSearchOptions(allResults: true, maxDepth: 1));

            Paths(outcome).Should().BeEquivalentTo("root", "a");
        }

        [Fact]
        public void A_match_deeper_than_the_limit_is_excluded()
        {
            SearchOver(DepthTree).Search("gamma", new FolderSearchOptions(maxDepth: 2))
                .Results.Should().BeEmpty();
            Paths(SearchOver(DepthTree).Search("gamma", new FolderSearchOptions(maxDepth: 3)))
                .Should().Equal("c");
        }

        [Fact]
        public void A_matching_ancestor_within_the_limit_still_collapses_its_deep_subtree()
        {
            var outcome = SearchOver(DepthTree).Search("alpha", new FolderSearchOptions(maxDepth: 1));

            var only = outcome.Results.Should().ContainSingle().Which;
            only.Folder.EntryId.Should().Be("a");
            only.IsCollapsed.Should().BeTrue();
        }

        [Fact]
        public void With_all_results_a_level_N_folder_with_children_carries_the_collapse_suffix()
        {
            var outcome = SearchOver(DepthTree).Search("alpha",
                new FolderSearchOptions(allResults: true, maxDepth: 2));

            outcome.Results.Select(r => r.Folder.EntryId).Should().Equal("a", "b");
            outcome.Results.Single(r => r.Folder.EntryId == "a").IsCollapsed.Should().BeFalse();
            outcome.Results.Single(r => r.Folder.EntryId == "b").IsCollapsed.Should().BeTrue();
        }

        [Fact]
        public void A_folder_name_containing_the_separator_gets_the_right_depth()
        {
            var nodes = new[]
            {
                new FolderNode("s1", "x", null, "A / B", "Archive / A / B", isLeaf: false),
                new FolderNode("s1", "y", "x", "Child", "Archive / A / B / Child", isLeaf: true),
            };

            Paths(SearchOver(nodes).Search("child", new FolderSearchOptions(maxDepth: 1)))
                .Should().BeEmpty();
            Paths(SearchOver(nodes).Search("child", new FolderSearchOptions(maxDepth: 2)))
                .Should().Equal("y");
        }

        [Fact]
        public void The_cap_counts_results_after_depth_filtering()
        {
            var outcome = SearchOver(DepthTree).Search("archive",
                new FolderSearchOptions(allResults: true, maxResults: 2, maxDepth: 2));

            outcome.TotalMatchCount.Should().Be(3); // root, a, b: c is too deep
            outcome.LimitExceeded.Should().BeTrue();
            outcome.Results.Should().HaveCount(2);
        }

        [Fact]
        public void A_parent_cycle_does_not_hang_the_depth_computation()
        {
            var nodes = new[]
            {
                new FolderNode("s1", "p", "q", "P", "Archive / P", isLeaf: false),
                new FolderNode("s1", "q", "p", "Q", "Archive / Q", isLeaf: false),
            };

            var outcome = SearchOver(nodes).Search("archive", new FolderSearchOptions(allResults: true, maxDepth: 5));

            outcome.Results.Should().HaveCount(2);
        }
    }
}
