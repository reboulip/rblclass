# RBLclass — Security architecture & threat model

> Status: living document. Baseline security audit **2026-06-07** (Phase 9)
> found **no critical or high-severity vulnerabilities**. This file records the
> trust boundaries, the data handled, and the security-relevant posture of the
> add-in so a corporate security review has an explicit trail. It is updated
> whenever a change alters that posture — most recently for **v2.10.0.0** D1
> (banner diagnostics artifact dump), which introduced the first on-demand
> write of mail body content to disk (see [Data handled](#data-handled)).

## What RBLclass is

A per-user Outlook **Shared COM Add-in** (.NET Framework 4.8, AnyCPU) that files
emails into folders across the user's mailboxes and PST archives, with fast
keyword search over the folder tree, attachment handling, and send-time guards.
It installs entirely under the current user (HKCU + `%LocalAppData%`) with **no
administrator rights** and runs **inside the Outlook process**.

## Trust boundaries

| Boundary | Description |
| --- | --- |
| **Outlook process ↔ add-in** | The add-in is loaded in-process by Outlook and runs with the user's full Outlook privileges. It is not a security boundary against the host; it trusts the Outlook Object Model and is trusted by it. |
| **User account ↔ add-in data** | All persisted state lives under `%LocalAppData%\RBLclass\` and is protected by the standard per-user NTFS ACL. Other non-admin users on the machine cannot read it; a local administrator can (accepted — the OS trust boundary). |
| **Machine ↔ network** | By default the add-in makes **no** network calls. The single optional exception is the opt-in update check (below), which reaches out to a user-configured HTTPS URL only. |
| **Decrypted mail ↔ add-in** | With the opt-in, experimental `ReencryptProtectedOnFile` setting (v2.12.0.0, off by default), the add-in reads mail content that Outlook and an endpoint product such as Stormshield have already decrypted in memory, and encrypts it again locally (see Protected mail). |
| **Add-in ↔ registry** | COM registration and the Outlook add-in load entry are written under **HKCU only** (`Software\Classes\...` and `Software\Microsoft\Office\Outlook\Addins\`). No HKLM writes; no admin rights required. |

## Data handled

- **Folder tree index (persisted):** folder display names and paths across the
  user's stores, in a SQLite database (`%LocalAppData%\RBLclass\`), with an FTS5
  index for search. **No mail bodies are persisted.** Subjects are read at
  classify time but are not stored in the index.
- **Settings (persisted):** user preferences in the SQLite `Settings` table,
  including the send-guard configuration (`SendExternalWarning`,
  `InternalDomains`, `ForgottenAttachmentKeywords`) and, for v2.7.0.0, the
  update-check toggle and URL.
- **Transient send-time metadata (not persisted):** recipient addresses and body
  text are inspected in memory by the send-time guards (forgotten-attachment and
  external-recipient checks) and discarded when the send completes. They are
  **not** written to disk or to the log.
- **Mail content (not persisted by classify operations):** classify operations
  move/copy items via the Outlook Object Model; mail bodies and attachments are
  handled by Outlook, never copied into RBLclass storage.
- **Banner diagnostic artifact (persisted on demand):** when the user explicitly
  clicks *"Save diagnostic file"* in **Settings → External sender banner**, the
  full HTML body of the currently selected mail is written to
  `%LocalAppData%\RBLclass\diagnostics\banner-diag-<timestamp>.html`, with any
  matched banner substring visually highlighted and a metadata comment (match
  outcome, signature length, timestamp). This is the **only** path that writes
  mail body content to disk. It is never triggered automatically; the
  `diagnostics\` directory is created lazily on first use; files accumulate until
  the user manually deletes them (no rolling retention, unlike the log files).
- **Re-encrypted filing (transient, opt-in, v2.12.0.0):** decrypted content
  (body, body images, attachments) is held in memory while a copy is rebuilt and
  encrypted; the plaintext MIME buffer is cleared after encryption. Certificates
  are read from `CurrentUser\My` (public key only). Temp-file plaintext: see
  Protected mail.

## Data at rest

- Location: `%LocalAppData%\RBLclass\` — SQLite database, rolling Serilog logs
  under `logs\`, and on-demand banner diagnostic HTML files under
  `diagnostics\`.
- Protection: standard per-user `%LocalAppData%` NTFS ACL — the correct OS trust
  boundary for non-secret, per-user data.
- **No application-level encryption** (accepted by design): the contents (folder
  names/paths, preferences) are non-secret; confidentiality relies on the
  per-user ACL. See [Accepted by design](#accepted-by-design).
- Installer log (outside `%LocalAppData%`): from version 2.11, every MSI
  install, upgrade and uninstall writes a verbose Windows Installer log to
  `%TEMP%\MSI*.LOG` (`MsiLogging` property in `installer/Package.wxs`). It
  contains local paths and the user name, no mail data, and is not cleaned up
  automatically.

## Log hygiene

- Sink: rolling file under `%LocalAppData%\RBLclass\logs\` (Serilog), per-user
  ACL, no remote sink.
- Policy: **no recipient addresses and no mail-body text reach the log.** The
  send-time guards log outcomes (e.g. "external-recipient warning shown"), not
  the addresses or body that triggered them. Folder names/paths and file/DB
  paths may appear for diagnostics. The re-encrypted filing path logs no subject,
  file names or addresses: only entry ids, thumbprint, algorithm OIDs, counts and
  fallback reasons.

## Protected (encrypted / signed) mail

- **One definition** (`RBLclass.Core/ProtectedMailPolicy.IsProtected`): a mail is
  protected when its message class starts with `IPM.Note.SMIME` or
  `IPM.Note.Secure` (ordinal, case-insensitive), or when its class is null/blank
  or cannot be read (fail closed). **The live class alone is not trusted**
  (GH #62): Stormshield decrypts the opened item in memory, so the opened
  `MailItem` reports a plain `IPM.Note` and a class-only check let the strip and
  its `Save()` through, writing the decrypted view without attachments. A mail is
  therefore also protected when any of these holds (`ProtectedMailPolicy.Evaluate`;
  `OutlookMailStore.IsProtectedMail`/`ReadProtection`):
  - the **stored** class, read from the parent folder's contents table
    (`Folder.GetTable`, restricted to protected classes, narrowed by
    `PR_INTERNET_MESSAGE_ID` and falling back to subject, row matched with
    `CompareEntryIDs`), is a protected class; a table row is what is stored and
    is unaffected by in-memory decryption;
  - the item still carries an S/MIME payload part named exactly `smime.p7m` or
    `smime.p7s` (other `.p7m` files, e.g. signed PDFs, are ordinary attachments);
  - the item already had unsaved changes when opened, before RBLclass touched it
    (something rewrote it in memory and a save would commit that); ignored for a
    draft being composed (reply/forward banner strip).
- **Never saved or rewritten.** Attachment removal, the HTML note, the banner
  strip, body-image removal, the disposition label, mark-as-read, and both
  task-flag saves (`MarkTaskComplete`,
  `MarkTaskIncomplete`) are guarded; the flag guard runs before the `FlagStatus`
  setter, which alone dirties the item. A flagged protected mail stays flagged
  and is excluded from the mark-complete prompt (`ClassifierService`); no new UI
  string.
- **Verdict taken once, on the source.** `ClassifierService` decides protection
  once on the source mail before filing (`IMailStore.IsEncryptedMail`,
  multi-signal) and does not re-derive it from the filed item, because a move can
  make the filed item look plain. When the source is protected, the filed copy
  (the move and every extra copy) gets no write at all: no attachment strip, body
  image removal, disposition label, banner strip, mark-as-read (`Move`/`Copy` are
  called with `markAsRead=false`) or task-flag completion. A requested strip
  counts as "encrypted kept" in the status. After filing, the filed item is
  checked again; if it no longer reads as protected, a Warning is always logged
  (not Debug-gated): "Protected mail filed by <move|copy> no longer reads as
  protected (<EntryId>): it was decrypted in memory when filed. No change was
  made to it." No subject is logged.
- **Known limitation: Move/Copy persist in-memory state.** Outlook's
  `MailItem.Move` and `Copy` write the item's in-memory state into the filed item
  while another handle holds that state modified; `Close(olDiscard)` on another
  handle does not prevent it, and once the holding handle releases the item a
  fresh handle sees the stored state again. Reproduced on the dev machine by
  simulation (stored `IPM.Note.SMIME`, class rewritten to `IPM.Note` in memory by
  a held handle: the stored class after the move was `IPM.Note`). If Stormshield
  has decrypted a mail in memory (smart card inserted, mail displayed), filing it
  therefore writes the decrypted version; this is the likely cause of the
  original #62 report ("the mail loses its encryption"). RBLclass cannot prevent
  this through the Outlook object model; by default the add-in does no
  re-encryption (see the experimental option below).
- **Experimental: re-encrypted filing (v2.12.0.0, GH #62, off by default).** With
  `ReencryptProtectedOnFile` on (Settings → Experimental), `ClassifierService`
  tries, for a protected mail that is readable decrypted in memory, a path that
  rebuilds it (`IMailRebuilder`, `OutlookMailRebuilder`), applies the classify
  options (`ReencryptContentPlanner`), encrypts it again and creates it directly
  as an `IPM.Note.SMIME` item in each destination, before falling back to the
  standard loop.
  - **Local, no network.** `EnvelopedCms` from the in-box `System.Security.dll`:
    CMS EnvelopedData, AES-256-CBC, RSA-OAEP key transport (the only RSA scheme
    .NET Framework 4.8 produces). Candidates come from `CurrentUser\My`, filtered
    on validity window, RSA key, Key Usage (key encipherment or none) and EKU
    (secure e-mail or none): no chain building, no revocation check, so no network
    access. Only the public key is used; RBLclass never uses the private key.
    *Automatic* (`ReencryptCertificateThumbprint` empty) additionally requires an
    e-mail address (subject e-mail or SAN rfc822Name) equal to the default
    account's SMTP address and a private key visible to Windows (so the copy is
    decryptable here); latest expiry wins. An explicit choice may be any
    candidate; if it stops being valid it is not silently replaced and
    re-encryption is skipped.
  - **Plaintext handling.** Attachment bytes are read through MAPI properties.
    When that fails, `Attachment.SaveAsFile` writes a plaintext temp file under
    `%LocalAppData%\RBLclass\tmp\<guid>`, which is read, overwritten with zeros,
    flushed and deleted (3 retries); leftovers are swept at startup. Zeroing is
    best effort (SSD wear levelling, EDR/AV copies). The MIME buffer is cleared
    after encryption; nothing decrypted is saved to a mail store by this path.
  - **Verification and fallback.** Each copy is verified as stored S/MIME (folder
    contents table only). A failure (no usable certificate, content not
    rebuildable, build/encryption, copy creation or verification) removes every
    copy made for that item and falls back to the standard path; the count is
    reported.
  - **The original.** Unless "keep a copy" is on, it goes to its store's Deleted
    Items. If that move wrote it decrypted (it no longer reads as protected), it
    is deleted permanently and a warning is logged; hard-deleted data may persist
    in PST free space until compaction. Mail already in Deleted Items is not
    moved.
  - **Not eligible** (standard path): signed-only mail, mail still showing
    `smime.p7m` (card absent), mail protected only by "unsaved changes",
    unresolved items. Not undoable; sender signature not kept; 25 MB cap.
- **Diagnostic "Protection trace"** (`ClassifierService`, via `IAppLog`): active
  only with Diagnostic logging on (effective after an Outlook restart). Debug
  lines record stage (before/after), path (pending, move, copy, safety-copy,
  kept-original), EntryID, message class, `PR_SECURITY_FLAGS` and the protected
  verdict; a warning is logged if a mail was protected before and is not after.
  No subject, body or recipient is logged. The `PropertyAccessor` read behind it
  may trigger Stormshield hooks, so it happens only with Diagnostic logging on.
  The trace also records the stored class, the payload and unsaved-on-open flags,
  and the deciding reason.
- **Non-class detection is logged.** When a signal other than the live class
  decides, an Information line `"<action>: protected mail detected by <Reason>
  (class=..., stored=..., entry=...)"` is written; no subject or body.
- **Unreadable selection diagnostics.** A selected item Outlook cannot read
  (exception or missing EntryID/StoreID) is logged as a Warning with the count
  and exception type/HRESULT; a classify that finds nothing to file logs an
  Information line. Cause of the "badge removed, Classify does nothing" symptom is
  not yet established.
- **`PR_SECURITY_FLAGS` (0x6E010003) is recorded only**, not used to decide
  protection, until validated on the real target.
- **Status: pending real-target verification.**
- **Body rewrite on explicit choice.** With the body-images setting on (off by
  default), removing a body image rewrites the mail's HTML body (each matching
  `<img src="cid:...">`, including Outlook's VML fallback, becomes a text
  placeholder) and deletes the attachment. This happens only on an explicit
  per-image user choice, never silently, is a local MAPI write with no network
  access, is skipped for protected mail (no body rewrite, no image removal), and
  cannot be undone.

## Network posture

**Default: no outbound network calls.** The add-in ships with no telemetry, no
auto-update download, and no analytics. This was the audit's finding at
2026-06-07 and remains the default. The **?** pane button and the About
window's *Online documentation* link (`DocumentationLink.TryOpen`) only launch
`https://reboulip.github.io/rblclass/` in the user's default browser on an
explicit click (`ShellExecute`); the add-in itself makes no request and sends
nothing.

**Single opt-in exception — the v2.7.0.0 update check (E1):**

- **Off by default, and inert without a URL.** It runs only when the user both
  enables *"Check for a newer version at startup"* **and** sets a
  *Version-check URL* in Settings → Updates. With either unset, no request is
  made (`ScheduleUpdateCheck` returns early).
- **What it does:** at Outlook startup, on a background thread, it issues **one
  anonymous HTTP GET** to the user-configured URL to read a plain-text version
  marker, compares it to the running assembly version, and — if newer — shows an
  orange in-pane banner. Dismissal snoozes the banner for 7 days.
- **What is sent:** only the GET request for the configured URL (standard
  `HttpClient` defaults). **No mail data, no folder names, no user identity, no
  telemetry** — nothing beyond the request for the version marker itself.
- **Proxy:** uses the shared static `HttpClient`, which honours the system /
  corporate proxy via `WebRequest.DefaultWebProxy`.
- **Robustness:** 5-second timeout, fully non-throwing — any failure (offline,
  proxy block, 404, malformed body) is logged at Information level and ignored;
  it never blocks startup or disrupts Outlook.
- **Recommended target:** a plain-text file on the organisation's **internal
  HTTPS share** (e.g. next to the install kit), so the check stays inside the
  corporate boundary.

Implementation: `RBLclass.AddIn/HttpUpdateCheckService.cs`,
`RBLclass.AddIn/RblClassAddIn.cs` (`ScheduleUpdateCheck`), decision logic in
`RBLclass.Core/UpdateCheck.cs`.

## Install & tamper surface

- **HKCU-only / no-admin:** install writes only current-user registry keys and
  `%LocalAppData%` files; uninstall removes them. No machine-wide footprint.
- **Signing:** Authenticode signing of the DLL and MSI with the internal-PKI
  Code Signing certificate (OID `1.3.6.1.5.5.7.3.3`), timestamped, is the GA
  target. Outlook does not require a signature to load the add-in (validated in
  Phase 0), so signing is currently optional.
- **Settings are user-writable.** The security-relevant guard settings live in
  the per-user SQLite `Settings` table. An optional **HKLM policy override**
  (GPO-pushable, admin wins) is a recorded Phase 9 hardening item, not yet
  implemented.

## Known hardening items (Phase 9 backlog)

Tracked in [ROADMAP.md](../ROADMAP.md) Phase 9. Highlights still open:

- Admin-lockable security settings via HKLM policy override.
- Friendly error dialogs (stop surfacing raw `ex.ToString()` to users; log the
  detail instead).
- Replace silent `catch { }` blocks with logged catches — especially the
  attachment-strip path, so a filed copy never silently keeps attachments the
  user believes were removed.
- Decide & document send-guard fail-open vs. fail-safe stance.
- CI: dependency vulnerability scanning, SAST, SBOM per release.
- Input caps: length-limit free-text inputs; upper-bound `MaxResults`; validate
  `InternalDomains` shape.
- Re-encrypted filing uses RSA-OAEP only; if Stormshield requires PKCS#1 v1.5, a
  `CryptMsg` P/Invoke follow-up is needed.

## Accepted by design

- **Plaintext SQLite settings / folder index** — contents are non-secret;
  confidentiality relies on the per-user `%LocalAppData%` ACL (the correct OS
  trust boundary). No app-level encryption planned.
- **External-recipient and diagnostics dialogs** that display addresses/paths
  are **intentional UX**, not information-disclosure leaks.
- **Send-guard fails open** — per the project rule that exceptions must never
  escape into Outlook (which would disable the add-in). A security team wanting
  fail-safe behaviour is addressed by the Phase 9 backlog item above.
- **Best-effort temp-file zeroing and PST free-space remnants** in the
  experimental re-encrypted filing path (see Protected mail); off by default.
