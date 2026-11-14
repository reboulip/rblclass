namespace RBLclass.Core
{
    /// <summary>
    /// What happened to one protected mail under re-encrypted filing (v2.12.0.0).
    /// The values after <see cref="Unresolved"/> are fallbacks: the mail was filed
    /// by the standard path instead, and they are counted and reported.
    /// </summary>
    public enum ReencryptOutcome
    {
        /// <summary>Rebuilt, re-encrypted and filed.</summary>
        Reencrypted = 0,

        /// <summary>Signed but not encrypted: nothing was decrypted, standard path (not a fallback).</summary>
        SignedOnly,

        /// <summary>Still shows its S/MIME payload (smart card absent): standard path (not a fallback).</summary>
        NotDecrypted,

        /// <summary>The item could not be resolved: standard path (not a fallback).</summary>
        Unresolved,

        /// <summary>No usable encryption certificate.</summary>
        NoCertificate,

        /// <summary>Content not rebuildable (embedded item, size cap, read failure, attachment map mismatch).</summary>
        Unsupported,

        /// <summary>Building the MIME entity failed.</summary>
        BuildFailed,

        /// <summary>Encryption failed.</summary>
        EncryptFailed,

        /// <summary>Creating an encrypted copy failed.</summary>
        CopyFailed,

        /// <summary>An encrypted copy did not verify as stored S/MIME.</summary>
        VerifyFailed,
    }
}
