using System;
using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Builds and searches the favourite-folder filesystem index (v2.4.0.0 F1):
    /// walks the user's chosen root directories (from the
    /// <see cref="SettingsKeys.AttachmentFavoriteFolders"/> setting) via an
    /// injected <see cref="IDirectoryScanner"/>, persists the expanded tree, and
    /// searches it with the same matcher as the Outlook folder search. No I/O of
    /// its own - the scan is injected and persistence is behind
    /// <see cref="IFavoriteFolderRepository"/>, keeping Core portable.
    /// </summary>
    public sealed class FavoriteFolderService
    {
        /// <summary>Sub-directory recursion cap (32-bit memory budget - paths only).</summary>
        public const int DefaultMaxDepth = 4;

        /// <summary>Hard cap on indexed directories across all roots.</summary>
        public const int DefaultMaxTotal = 2000;

        private readonly IDirectoryScanner _scanner;
        private readonly IFavoriteFolderRepository _repo;
        private readonly ISettingsStore _settings;
        private readonly object _gate = new object();
        private IReadOnlyList<FavoriteFolder> _cache = new FavoriteFolder[0];

        public FavoriteFolderService(IDirectoryScanner scanner,
                                     IFavoriteFolderRepository repo,
                                     ISettingsStore settings)
        {
            _scanner = scanner ?? throw new ArgumentNullException(nameof(scanner));
            _repo = repo ?? throw new ArgumentNullException(nameof(repo));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>Load the persisted index into the search cache (startup, no filesystem walk).</summary>
        public void LoadFromCache()
        {
            var all = _repo.LoadFavorites();
            lock (_gate) { _cache = all; }
        }

        /// <summary>
        /// Re-walk every favourite root from settings, persist the expanded tree,
        /// and refresh the search cache. May run on a background thread - it
        /// touches the filesystem (via the scanner) and SQLite, never Outlook COM.
        /// </summary>
        public void Reindex()
        {
            var roots = Settings.ParseList(
                _settings.Get(SettingsKeys.AttachmentFavoriteFolders, string.Empty));

            var all = new List<FavoriteFolder>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var root in roots)
            {
                if (all.Count >= DefaultMaxTotal) break;
                if (string.IsNullOrWhiteSpace(root) || !seen.Add(root)) continue;

                int budget = DefaultMaxTotal - all.Count;
                foreach (var dir in _scanner.GetDirectories(root, DefaultMaxDepth, budget))
                {
                    if (all.Count >= DefaultMaxTotal) break;
                    if (dir != root && !seen.Add(dir)) continue; // root already in seen
                    all.Add(new FavoriteFolder(dir, LastSegment(dir)));
                }
            }

            _repo.SaveFavorites(all);
            lock (_gate) { _cache = all.ToArray(); }
        }

        /// <summary>
        /// Search the favourite index with the shared folder-search matcher
        /// (AND across keywords, case/accent-insensitive). Returns
        /// <see cref="FolderSearchOutcome"/> whose results carry the directory
        /// path as <see cref="FolderNode.FullPath"/>.
        /// </summary>
        public FolderSearchOutcome Search(string query, FolderSearchOptions options = null)
        {
            IReadOnlyList<FavoriteFolder> snapshot;
            lock (_gate) { snapshot = _cache; }

            return new FolderSearchService(new FavoriteFolderSearchAdapter(snapshot))
                .Search(query, options);
        }

        /// <summary>Last path segment, e.g. "2024" from "C:\Archive\Contracts\2024".</summary>
        private static string LastSegment(string path)
        {
            if (string.IsNullOrEmpty(path)) return path;
            string trimmed = path.TrimEnd('\\', '/');
            int idx = trimmed.LastIndexOfAny(new[] { '\\', '/' });
            return idx >= 0 && idx < trimmed.Length - 1 ? trimmed.Substring(idx + 1) : trimmed;
        }
    }
}
