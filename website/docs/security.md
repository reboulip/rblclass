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
- **Mail content (never persisted):** classify operations move/copy items via the
  Outlook Object Model; bodies and attachments are handled by Outlook, never
  copied into RBLclass storage.

## Data at rest

- **Location:** `%LocalAppData%\RBLclass\` (SQLite database + rolling Serilog
  logs).
- **Protection:** the standard per-user `%LocalAppData%` NTFS ACL — the correct
  OS trust boundary for non-secret, per-user data.
- **No application-level encryption** (accepted by design): the contents (folder
  names/paths, preferences) are non-secret; confidentiality relies on the
  per-user ACL.

## Log hygiene

- **Sink:** rolling file under `%LocalAppData%\RBLclass\logs\` (Serilog),
  per-user ACL, **no remote sink**.
- **Policy:** no recipient addresses and no mail-body text reach the log. The
  send-time guards log *outcomes* (e.g. "external-recipient warning shown"), not
  the addresses or body that triggered them. Folder names/paths and file/DB paths
  may appear for diagnostics.

## Network posture

**Default: no outbound network calls.** The add-in ships with no telemetry, no
auto-update download, and no analytics.

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
- Input caps: length-limit free-text inputs; upper-bound `MaxResults`; validate
  `InternalDomains` shape.
