# Security model

!!! abstract "Status"
    Living document. The baseline security audit of **2026-06-07** (Phase 9)
    found **no critical or high-severity vulnerabilities**. This page records the
    trust boundaries, the data handled, and the security-relevant posture of the
    add-in so a corporate security review has an explicit trail. It is updated
    whenever a change alters that posture.

    The repo also keeps a fuller internal threat model at
    [`docs/security.md`](https://github.com/reboulip/rblclass/blob/main/docs/security.md);
    the two are kept in sync.

## What RBLclass is

A per-user Outlook **Shared COM Add-in** (.NET Framework 4.8, AnyCPU) that files
emails into folders across the user's mailboxes and PST archives, with fast
keyword search over the folder tree, attachment handling, and send-time guards.
It installs entirely under the current user (HKCU + `%LocalAppData%`) with **no
administrator rights** and runs **inside the Outlook process**.

## Trust boundaries

| Boundary | Description |
| --- | --- |
| **Outlook process ↔ add-in** | The add-in is loaded in-process by Outlook and runs with the user's full Outlook privileges. It is **not** a security boundary against the host: it trusts the Outlook Object Model and is trusted by it. |
| **User account ↔ add-in data** | All persisted state lives under `%LocalAppData%\RBLclass\`, protected by the standard per-user NTFS ACL. Other non-admin users cannot read it; a local administrator can (accepted — the OS trust boundary). |
| **Machine ↔ network** | By default the add-in makes **no** network calls. The single optional exception is the opt-in update check, which reaches a user-configured HTTPS URL only. |
| **Decrypted mail ↔ add-in** | With the opt-in, experimental *Re-encrypt encrypted mail when filing* setting (off by default), the add-in reads mail content that Outlook and an endpoint product such as Stormshield have already decrypted in memory, and encrypts it again locally. See [Protected mail](#protected-encrypted-signed-mail). |
| **Add-in ↔ Microsoft Word** | The optional *Save as file → PDF* export drives Word through late-bound automation, **out of process**, to convert mail to PDF. Word is optional at runtime (no Word interop reference, no new dependency); without it, nothing is written. Word runs on its own dedicated STA thread; Outlook OM calls stay on the UI thread. |
| **Add-in ↔ registry** | COM registration and the Outlook load entry are written under **HKCU only**. No HKLM writes; no admin rights. |

## Data handled

- **Folder tree index (persisted):** folder display names and paths across the
  user's stores, in a SQLite database with an FTS5 search index. **No mail bodies
  are persisted.** Subjects are read at classify time but are not stored.
- **Settings (persisted):** user preferences in the SQLite `Settings` table,
  including the send-guard configuration and the update-check toggle/URL.
- **Transient send-time metadata (not persisted):** recipient addresses and body
  text are inspected in memory by the send-time guards and discarded when the
  send completes. They are **not** written to disk or to the log.
- **Mail content (not persisted by classify operations):** classify operations
  move/copy items via the Outlook Object Model; bodies and attachments are handled
  by Outlook, never copied into RBLclass storage.
- **Banner diagnostic artifact (persisted on demand):** when the user explicitly
  clicks *"Save diagnostic file"* in **Settings → External sender banner**, the
  full HTML body of the currently selected mail is written to
  `%LocalAppData%\RBLclass\diagnostics\banner-diag-<timestamp>.html`, with any
  matched banner substring visually highlighted and a metadata comment. This is
  the **only** path that writes mail body content to disk. It is never triggered
  automatically; the `diagnostics\` directory is created on first use; files
  accumulate until the user manually deletes them (no rolling retention).
- **External-attachment phishing warning (transient, on by default):** the
  add-in hooks two Outlook item events (`BeforeAttachmentWriteToTempFile`,
  `BeforeAttachmentPreview`) on the selected mail and on opened mail windows. The
  mail's HTML body is read **only when the event fires**, to look for the learned
  banner, and is not persisted. Nothing is sent over the network and no mail is
  modified. The only output is a confirmation dialog (default and Esc: cancel);
  if the check fails, it fails open and the attachment opens. The log gets only
  counts and outcomes, never file names or addresses.
- **External-link phishing warning (transient, on by default):** the add-in
  installs two **thread-scoped** Windows hooks (`WH_MOUSE`, `WH_GETMESSAGE`) on
  Outlook's own UI thread, never global, with nothing injected and no API
  patched, and repurposes the built-in ribbon command `HyperlinkOpen`. When a link
  is activated in a mail, the mail's **rendered text** is read at click time to
  look for the learned banner; nothing is persisted, nothing goes over the
  network and no mail is modified. The only output is a confirmation dialog
  (default and Esc: cancel). It fails open: on any error the link opens, and
  "Open anyway" goes through Word's own hyperlink follow, so Outlook's security
  prompts still apply. The log gets only `scheme://host`, never the full URL.
  `mailto:` and `#` links are not checked. **Known gap:** a touch tap cannot be
  blocked (accepted, documented).
- **PDF export (user-initiated, plaintext):** *Save as file → PDF* writes each
  mail as a temporary MHTML file under `%TEMP%\RBLclass\pdf\`, which Word converts
  to a PDF in the folder the user picked. The temporary files hold plaintext mail
  and are always deleted, in a `finally` block. A PDF of **encrypted** mail is
  created only after an explicit confirmation for the whole batch (default and
  Esc: skip encrypted mails). No mail content reaches the log.
- **Re-encrypted filing (transient, opt-in):** decrypted content (body, body
  images, attachments) is held in memory while a copy is rebuilt and encrypted;
  the plaintext MIME buffer is cleared after encryption. Certificates are read
  from the user's `CurrentUser\My` store (public key only). See *Protected
  (encrypted / signed) mail*.

## Data at rest

- **Location:** `%LocalAppData%\RBLclass\` — SQLite database, rolling Serilog
  logs under `logs\`, and on-demand banner diagnostic HTML files under
  `diagnostics\`.
- **Protection:** the standard per-user `%LocalAppData%` NTFS ACL — the correct
  OS trust boundary for non-secret, per-user data.
- **No application-level encryption** (accepted by design): the contents (folder
  names/paths, preferences) are non-secret; confidentiality relies on the
  per-user ACL.
- **Installer log (outside `%LocalAppData%`):** from version 2.11, every MSI
  install, upgrade and uninstall writes a verbose Windows Installer log to
  `%TEMP%\MSI*.LOG`. It contains local paths and the user name, **no mail data**,
  and is **not cleaned up automatically**.

## Log hygiene

- **Sink:** rolling file under `%LocalAppData%\RBLclass\logs\` (Serilog),
  per-user ACL, **no remote sink**.
- **Policy:** no recipient addresses and no mail-body text reach the log. The
  send-time guards log *outcomes* (e.g. "external-recipient warning shown"), not
  the addresses or body that triggered them. Folder names/paths and file/DB paths
  may appear for diagnostics. The re-encrypted filing path logs no subject, file
  names or addresses: only entry ids, certificate thumbprint, algorithm OIDs,
  counts and fallback reasons.

## Protected (encrypted / signed) mail

- **One definition.** A mail is *protected* when its message class starts with
  `IPM.Note.SMIME` or `IPM.Note.Secure` (case-insensitive), **or when its class
  cannot be read** (fail closed). The live class alone is not trusted, because
  Stormshield decrypts an opened mail in memory and it then reports a plain
  class. A mail is therefore also protected when its **stored** class (read from
  the folder's contents table) is protected, when it still carries an S/MIME
  payload part named `smime.p7m` or `smime.p7s` (other `.p7m` files such as
  signed PDFs are ordinary attachments), or when it already had unsaved changes
  as soon as it was opened, before the add-in touched it (ignored for a draft
  being composed).
- **Never saved or rewritten.** For protected mail the add-in skips attachment
  removal, body-image removal, the disposition label, the HTML note, the banner
  strip, mark-as-read, and both task-flag
  saves (mark complete / incomplete). A flagged protected mail stays flagged and
  is not offered in the mark-complete prompt.
- **Decided once, on the source.** Protection is decided on the source mail
  before filing and is not re-derived from the filed item. When the source is
  protected, the filed copy (the move and every extra copy) gets no write at all:
  no attachment or body-image removal, label, banner strip, mark-as-read or task
  completion. A requested strip is reported as kept. If the filed item no longer
  reads as protected, a warning is always logged (no subject).
- **Cause of the "mail loses its encryption" report.** Outlook's move and copy
  write the item's in-memory state into the filed item. If an endpoint product
  such as Stormshield has decrypted a mail in memory (smart card inserted, mail
  displayed), filing it through the object model writes the decrypted version.
- **Store-level filing (closes the problem above on every filing path).**
  Protected mail is moved, copied and deleted at store level through Extended
  MAPI and is never opened through the Outlook object model when filed, so no
  decrypted state is written. This covers move, copy, keep a copy,
  multi-destination, safety copy, AutoClass, Undo, sent-item triage and delete.
    - **No fallback.** If the transfer fails, the mail stays where it was, still
      encrypted, and is reported as an error; it is never filed through the
      object model.
    - **Rare case.** If the filed mail's new position cannot be identified, it is
      filed but cannot be undone and is not used for a safety copy.
    - **Surface.** No network call, no new data persisted; the failure is logged
      with its error code and no subject.
    - **Status:** pending verification on the real target workstation (smart card
      inserted, 32-bit Outlook).
- **Experimental: re-encrypted filing (off by default).** With *Re-encrypt
  encrypted mail when filing* on (Settings → Experimental), an encrypted mail that
  is readable decrypted in memory is rebuilt from its content, encrypted again and
  created directly as an S/MIME item in each destination, instead of letting
  Outlook persist the decrypted version. It only runs when the filing changes the
  mail's content (attachments or banner removed, or an attachment or body image
  removed or saved out); otherwise the mail is filed as it is, keeps its original
  encryption and signature, and can be undone.
    - **Local, no network.** The add-in builds the CMS EnvelopedData itself, in
      managed code with no new dependency (AES-256-CBC, RSA PKCS#1 v1.5 key
      transport, one recipient: the same shape as a real Stormshield mail). The
      in-box `System.Security.dll` is only used to decode and check the result.
      Acceptance by Stormshield is pending on the real target. No
      network call is made. Certificates come from `CurrentUser\My`, filtered on
      validity window, RSA key, key usage and EKU only: no chain building and no
      revocation check. Only the public key is used; RBLclass never uses the
      private key. *Automatic* selection also requires the certificate's e-mail to
      equal the default account's SMTP address and a private key visible to Windows,
      so the copy is decryptable on this machine. An explicit choice that becomes
      invalid is not silently replaced: re-encryption is skipped.
    - **Plaintext handling.** Attachment bytes are read through MAPI properties.
      When that fails, `Attachment.SaveAsFile` writes a plaintext temporary file
      under `%LocalAppData%\RBLclass\tmp\<guid>`, which is read, overwritten with
      zeros, flushed and deleted (3 retries); leftovers are swept at startup.
      Zeroing is best effort: SSD wear levelling and EDR/antivirus copies can keep
      data. Nothing decrypted is saved to a mail store by this path.
    - **Verification and fallback.** Each copy is verified as stored S/MIME. Any
      failure (no usable certificate, content not rebuildable, encryption, copy or
      verification failure) falls back to the standard path after removing every
      copy made for that mail; the count is reported to the user.
    - **The original.** Unless *Keep a copy* is on, the original goes to Deleted
      Items. If Outlook's move wrote it decrypted, it is deleted permanently and a
      warning is logged; deleted data may persist in PST free space until
      compaction.
    - **Limits.** Not undoable; the sender's signature is not kept; signed-only
      mail, mail still showing its encrypted part and mail not readable decrypted
      are not eligible; 25 MB attachment cap.
- **Diagnostic "Protection trace".** Only when *Diagnostic logging* is on (takes
  effect after restarting Outlook), the log records, per classified mail, the
  stage (before/after), path (pending, move, copy, safety-copy, kept-original),
  EntryID, message class, `PR_SECURITY_FLAGS` value and whether the mail counts as
  protected, plus a warning if a mail was protected before and is not after. No
  subject, body or recipient is logged. The trace reads `PropertyAccessor`, which
  may trigger endpoint-security hooks, so that read happens only with Diagnostic
  logging on. The trace also records the stored class, the payload and
  unsaved-on-open flags, and the reason a mail counted as protected.
- **Non-class detection is logged.** When a signal other than the live class
  decides, an Information line records the action, reason, classes and EntryID;
  no subject or body.
- **Re-encryption step breadcrumbs.** The experimental re-encrypted filing logs an
  Information line before each step, so a crash can be pinpointed without
  Diagnostic logging. Each line holds only a step name and an EntryID: no subject,
  body, address or file name.
- **Unreadable selection diagnostics.** A selected item Outlook cannot read is
  logged as a Warning (count, exception type/HRESULT), and a classify that finds
  nothing to file logs an Information line.
- **`PR_SECURITY_FLAGS` is recorded only**, not used to decide protection, until
  validated on the real target.
- **Status:** pending verification on the real target workstation.
- **Body rewrite on explicit choice.** With *Offer body images for removal* on,
  removing a body image rewrites the mail's HTML body (each matching
  `<img src="cid:...">` becomes a text placeholder) and deletes the attachment.
  This only happens on an explicit per-image user choice, is a local MAPI write
  with no network access, is never applied to protected mail, and cannot be
  undone.

## Network posture

**Default: no outbound network calls.** The add-in ships with no telemetry, no
auto-update download, and no analytics. The **?** pane button and the About
window's *Online documentation* link only ask your default browser to open the
public documentation site on an explicit click; the add-in itself makes no
request and sends nothing.

The single opt-in exception is the **update check**:

- **Off by default, and inert without a URL.** It runs only when the user both
  enables *"Check for a newer version at startup"* **and** sets a *Version-check
  URL* in **Settings → Updates**. With either unset, no request is made.
- **What it does:** at Outlook startup, on a background thread, it issues **one
  anonymous HTTP GET** to the configured URL to read a plain-text version marker,
  compares it to the running assembly version, and — if newer — shows an orange
  in-pane banner. Dismissal snoozes the banner for 7 days.
- **What is sent:** only the GET request for the configured URL. **No mail data,
  no folder names, no user identity, no telemetry** — nothing beyond the request
  for the version marker.
- **Proxy:** uses the shared static `HttpClient`, honouring the system/corporate
  proxy via `WebRequest.DefaultWebProxy`.
- **Robustness:** 5-second timeout, fully non-throwing — any failure (offline,
  proxy block, 404, malformed body) is logged and ignored; it never blocks
  startup or disrupts Outlook.
- **Recommended target:** a plain-text file on the organisation's **internal
  HTTPS share**, so the check stays inside the corporate boundary.

## Install & tamper surface

- **HKCU-only / no-admin:** install writes only current-user registry keys and
  `%LocalAppData%` files; uninstall removes them. No machine-wide footprint.
- **Signing:** Authenticode signing of the DLL and MSI with the internal-PKI Code
  Signing certificate (EKU `1.3.6.1.5.5.7.3.3`), timestamped, is the GA target.
  Outlook does not require a signature to load the add-in, so signing is currently
  optional.
- **Settings are user-writable.** The security-relevant guard settings live in the
  per-user SQLite `Settings` table. An optional **HKLM policy override**
  (GPO-pushable, admin wins) is a recorded hardening item, not yet implemented.

## Accepted by design

- **Plaintext SQLite settings / folder index** — contents are non-secret;
  confidentiality relies on the per-user `%LocalAppData%` ACL. No app-level
  encryption planned.
- **External-recipient and diagnostics dialogs** that display addresses/paths are
  **intentional UX**, not information-disclosure leaks.
- **Best-effort temp-file zeroing and PST free-space remnants** in the
  experimental re-encrypted filing path (see *Protected (encrypted / signed)
  mail*); the feature is off by default.
- **Send-guard fails open** — per the project rule that exceptions must never
  escape into Outlook (which would disable the add-in). A fail-safe stance is a
  tracked hardening item.

## Known hardening items (Phase 9 backlog)

Tracked in
[`ROADMAP.md`](https://github.com/reboulip/rblclass/blob/main/ROADMAP.md) Phase 9.
Highlights still open:

- Admin-lockable security settings via HKLM policy override.
- Friendly error dialogs (stop surfacing raw exception text to users).
- Replace silent `catch { }` blocks with logged catches — especially the
  attachment-strip path, so a filed copy never silently keeps attachments the
  user believes were removed.
- Decide & document the send-guard fail-open vs. fail-safe stance.
- CI: dependency vulnerability scanning, SAST, SBOM per release.
- Re-encrypted filing now uses RSA PKCS#1 v1.5 key transport, the scheme a real
  Stormshield mail uses; acceptance by Stormshield with the smart card is still to
  be verified on the real target.
- Input caps: length-limit free-text inputs; upper-bound `MaxResults`; validate
  `InternalDomains` shape.
