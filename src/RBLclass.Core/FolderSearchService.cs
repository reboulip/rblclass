using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// In-memory <see cref="IFolderSearch"/> over the cached folder tree from
    /// <see cref="IFolderTree"/>. Reproduces the legacy <c>kwSearch</c>:
    /// <list type="bullet">
    ///   <item>A folder "matches" when ALL keywords are found on its full path
    ///   (the legacy <c>up</c> condition). Because a matching folder's whole
    ///   subtree also matches, results are the <b>topmost</b> matches; each
    ///   matched non-leaf is collapsed to "path | [...]" unless All results is
    ///   on, in which case every matching folder is listed.</item>
    ///   <item>Keyword test is AND, case/accent insensitive, per
    ///   <see cref="FolderMatchMode"/> (word-prefix default, substring opt-in).</item>
    ///   <item>Results are sorted alphabetically and capped at
    ///   <see cref="FolderSearchOptions.MaxResults"/>.</item>
    /// </list>
    /// </summary>
    public sealed class FolderSearchService : IFolderSearch
    {
        private readonly IFolderTree _tree;

        // Normalized-path cache, rebuilt only when the tree's CacheVersion
        // changes, so Normalize()/SplitWords() aren't re-run per folder on every
        // keystroke-triggered search (v2.7.1.0 E3).
        private Dictionary<string, NormalizedFolder> _normCache =
            new Dictionary<string, NormalizedFolder>();
        private int _lastCacheVersion = -1;

        // Folder depth by (StoreId, EntryId), rebuilt with the normalized-path cache
        // (v2.13.0.0 B1). Store root = 0, its direct sub-folders = 1.
        private Dictionary<string, int> _depthCache = new Dictionary<string, int>();

        public FolderSearchService(IFolderTree tree)
        {
            _tree = tree ?? throw new ArgumentNullException(nameof(tree));
        }

        /// <summary>
        /// Precomputed normalized forms of a folder's <see cref="FolderNode.FullPath"/>,
        /// cached between searches (v2.7.1.0 E3). FullPath is immutable between
        /// reindexes, so this is computed once per folder per tree generation.
        /// </summary>
        private sealed class NormalizedFolder
        {
            public readonly string PathNoSpaces;
            public readonly List<string> Words;

            public NormalizedFolder(string fullPath)
            {
                string normalizedPath = TextNormalization.Normalize(fullPath);
                PathNoSpaces = RemoveWhitespace(normalizedPath);
                Words = SplitWords(normalizedPath);
            }
        }

        public FolderSearchOutcome Search(string query, FolderSearchOptions options = null)
        {
            options = options ?? FolderSearchOptions.Default;

            // Too short to search yet (v2.2 responsiveness fix - one or two
            // keystrokes used to match most of the tree).
            if ((query ?? string.Empty).Trim().Length < options.MinQueryLength)
                return FolderSearchOutcome.Empty;

            var keywords = Tokenize(query);
            if (keywords.Count == 0)
                return FolderSearchOutcome.Empty;

            var all = _tree.GetAll();

            // Rebuild the normalized-path cache only when the tree changed.
            int treeVersion = _tree.CacheVersion;
            if (treeVersion != _lastCacheVersion)
            {
                _normCache = BuildNormCache(all);
                _depthCache = BuildDepthCache(all);
                _lastCacheVersion = treeVersion;
            }

            // 1) Every folder whose full path satisfies all keywords.
            var matches = new List<FolderNode>();
            foreach (var folder in all)
            {
                NormalizedFolder norm;
                if (!_normCache.TryGetValue(KeyOf(folder), out norm))
                    norm = new NormalizedFolder(folder.FullPath); // cache miss: compute once
                if (options.MaxDepth > 0 && DepthOf(folder) > options.MaxDepth)
                    continue; // deeper than the configured maximum search depth
                if (IsMatch(norm, keywords, options.MatchMode))
                    matches.Add(folder);
            }

            // 2) Collapse to topmost matches (unless All results expands them).
            List<FolderSearchResult> emitted;
            if (options.AllResults)
            {
                emitted = matches
                    // A level-N folder that still has sub-folders carries the collapse
                    // suffix so the user can tell its subtree was cut off.
                    .Select(f => new FolderSearchResult(f, isCollapsed:
                        options.MaxDepth > 0 && !f.IsLeaf && DepthOf(f) == options.MaxDepth))
                    .ToList();
            }
            else
            {
                var matchedKeys = new HashSet<string>(matches.Select(KeyOf));
                emitted = new List<FolderSearchResult>();
                foreach (var f in matches)
                {
                    bool parentMatched = f.ParentEntryId != null
                        && matchedKeys.Contains(Key(f.StoreId, f.ParentEntryId));
                    if (parentMatched)
                        continue; // subsumed by a matching ancestor

                    emitted.Add(new FolderSearchResult(f, isCollapsed: !f.IsLeaf));
                }
            }

            // 3) Alphabetical ordering (legacy sorted the result list).
            emitted.Sort((a, b) => string.Compare(
                a.Folder.FullPath, b.Folder.FullPath,
                StringComparison.InvariantCultureIgnoreCase));

            // 4) Cap.
            int total = emitted.Count;
            bool exceeded = total > options.MaxResults;
            IReadOnlyList<FolderSearchResult> results = exceeded
                ? emitted.Take(options.MaxResults).ToList()
                : emitted;

            return new FolderSearchOutcome(results, total, exceeded);
        }

        private int DepthOf(FolderNode folder)
        {
            int depth;
            return _depthCache.TryGetValue(KeyOf(folder), out depth) ? depth : 1;
        }

        /// <summary>
        /// Depth of every folder from the ParentEntryId chain (never by splitting
        /// FullPath: folder names may contain the separator). The adapter gives the
        /// store root node and the top-level folders the same null parent, so a null
        /// parent is level 0 only for the root node itself (its path is its name) and
        /// level 1 otherwise. A missing parent or a cycle ends the chain safely.
        /// </summary>
        private static Dictionary<string, int> BuildDepthCache(IEnumerable<FolderNode> all)
        {
            var byKey = new Dictionary<string, FolderNode>();
            foreach (var f in all) byKey[KeyOf(f)] = f;

            var depths = new Dictionary<string, int>(byKey.Count);
            foreach (var f in byKey.Values)
            {
                var chain = new List<string>();
                var seen = new HashSet<string>();
                var cur = f;
                int baseDepth;
                while (true)
                {
                    string key = KeyOf(cur);
                    int known;
                    if (depths.TryGetValue(key, out known)) { baseDepth = known; break; }
                    if (!seen.Add(key)) { baseDepth = 0; break; } // cycle guard
                    chain.Add(key);
                    if (cur.ParentEntryId == null)
                    {
                        baseDepth = string.Equals(cur.FullPath, cur.Name, StringComparison.Ordinal) ? -1 : 0;
                        // -1 + 1 = 0 for the root node; 0 + 1 = 1 for top-level folders.
                        break;
                    }
                    FolderNode parent;
                    if (!byKey.TryGetValue(Key(cur.StoreId, cur.ParentEntryId), out parent))
                    {
                        baseDepth = 0; // unknown parent: treat as a top-level folder
                        break;
                    }
                    cur = parent;
                }
                // chain is child-first; the last entry is the one whose base we know.
                int d = baseDepth;
                for (int i = chain.Count - 1; i >= 0; i--)
                {
                    d++;
                    depths[chain[i]] = d;
                }
            }
            return depths;
        }

        private static List<string> Tokenize(string query)
        {
            var tokens = new List<string>();
            if (string.IsNullOrWhiteSpace(query))
                return tokens;

            foreach (var raw in query.Split((char[])null, StringSplitOptions.RemoveEmptyEntries))
            {
                string norm = TextNormalization.Normalize(raw);
                if (norm.Length > 0)
                    tokens.Add(norm);
            }
            return tokens;
        }

        private static Dictionary<string, NormalizedFolder> BuildNormCache(
            IReadOnlyList<FolderNode> all)
        {
            var dict = new Dictionary<string, NormalizedFolder>(all.Count);
            foreach (var f in all)
                dict[KeyOf(f)] = new NormalizedFolder(f.FullPath);
            return dict;
        }

        private static bool IsMatch(NormalizedFolder norm, List<string> keywords, FolderMatchMode mode)
        {
            if (mode == FolderMatchMode.Substring)
            {
                foreach (var kw in keywords)
                {
                    if (norm.PathNoSpaces.IndexOf(kw, StringComparison.Ordinal) < 0)
                        return false;
                }
                return true;
            }

            // WordPrefix: each keyword must start some word of the path. Words
            // are letter/digit runs, so a keyword carrying any other character
            // (e.g. "r&d") could never start one - those keywords are tested by
            // containment instead, so "R&D" still finds the folder named "R&D".
            foreach (var kw in keywords)
            {
                bool any = false;
                if (HasNonWordChar(kw))
                {
                    any = norm.PathNoSpaces.IndexOf(kw, StringComparison.Ordinal) >= 0;
                }
                else
                {
                    foreach (var w in norm.Words)
                    {
                        if (w.StartsWith(kw, StringComparison.Ordinal)) { any = true; break; }
                    }
                }
                if (!any) return false;
            }
            return true;
        }

        private static bool HasNonWordChar(string keyword)
        {
            foreach (char ch in keyword)
            {
                if (!char.IsLetterOrDigit(ch))
                    return true;
            }
            return false;
        }

        private static List<string> SplitWords(string normalizedText)
        {
            var words = new List<string>();
            int start = -1;
            for (int i = 0; i < normalizedText.Length; i++)
            {
                if (char.IsLetterOrDigit(normalizedText[i]))
                {
                    if (start < 0) start = i;
                }
                else if (start >= 0)
                {
                    words.Add(normalizedText.Substring(start, i - start));
                    start = -1;
                }
            }
            if (start >= 0)
                words.Add(normalizedText.Substring(start));
            return words;
        }

        private static string RemoveWhitespace(string value)
        {
            var sb = new System.Text.StringBuilder(value.Length);
            foreach (char ch in value)
            {
                if (!char.IsWhiteSpace(ch))
                    sb.Append(ch);
            }
            return sb.ToString();
        }

        private static string KeyOf(FolderNode f) => Key(f.StoreId, f.EntryId);

        private static string Key(string storeId, string entryId) => storeId + " " + entryId;
    }
}
