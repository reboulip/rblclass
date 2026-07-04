using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Key/value settings persistence (the SQLite <c>Settings</c> table from
    /// schema v1). Replaces the legacy positional <c>options.txt</c>
    /// (deviation #3). Step 3 uses it for the task-pane toggles; Step 9 builds
    /// the full settings pane over the same store.
    /// </summary>
    public interface ISettingsStore
    {
        /// <summary>Create the Settings table if missing (idempotent).</summary>
        void EnsureSchema();

        /// <summary>Get a string value, or <paramref name="fallback"/> if absent.</summary>
        string Get(string key, string fallback = null);

        /// <summary>Set (insert or replace) a string value.</summary>
        void Set(string key, string value);

        /// <summary>Get a bool ("1"/"0"), or <paramref name="fallback"/> if absent/unparseable.</summary>
        bool GetBool(string key, bool fallback);

        /// <summary>Set a bool as "1"/"0".</summary>
        void SetBool(string key, bool value);

        /// <summary>
        /// Read every stored key in a single query. Keys absent from the store
        /// are absent from the result. Used by <see cref="Settings.Load"/> so a
        /// full load is one connection, not one per key (v2.7.1.0 E2).
        /// </summary>
        IDictionary<string, string> GetAll();

        /// <summary>
        /// Insert-or-replace every entry in a single transaction. Used by
        /// <see cref="Settings.Save"/> so a full save is one connection
        /// (v2.7.1.0 E2).
        /// </summary>
        void SetMany(IEnumerable<KeyValuePair<string, string>> entries);
    }
}
