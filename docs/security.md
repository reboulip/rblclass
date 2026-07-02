# RBLclass — Security architecture & threat model

> Status: living document. Baseline security audit **2026-06-07** (Phase 9)
> found **no critical or high-severity vulnerabilities**. This file records the
> trust boundaries, the data handled, and the security-relevant posture of the
> add-in so a corporate security review has an explicit trail. It is updated
> whenever a change alters that posture — most recently for the **v2.7.0.0**
> opt-in update check (E1), which introduced the product's first outbound
> network call (see [Network posture](#network-posture)).

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
- **Mail content (never persisted):** classify operations move/copy items via the
  Outlook Object Model; mail bodies and attachments are handled by Outlook, never
  copied into RBLclass storage.

## Data at rest

- Location: `%LocalAppData%\RBLclass\` (SQLite database + rolling Serilog logs).
- Protection: standard per-user `%LocalAppData%` NTFS ACL — the correct OS trust
  boundary for non-secret, per-user data.
- **No application-level encryption** (accepted by design): the contents (folder
  names/paths, preferences) are non-secret; confidentiality relies on the
  per-user ACL. See [Accepted by design](#accepted-by-design).

## Log hygiene

- Sink: rolling file under `%LocalAppData%\RBLclass\logs\` (Serilog), per-user
  ACL, no remote sink.
- Policy: **no recipient addresses and no mail-body text reach the log.** The
  send-time guards log outcomes (e.g. "external-recipient warning shown"), not
  the addresses or body that triggered them. Folder names/paths and file/DB
  paths may appear for diagnostics.

## Network posture

**Default: no outbound network calls.** The add-in ships with no telemetry, no
auto-update download, and no analytics. This was the audit's finding at
2026-06-07 and remains the default.

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

## Accepted by design

- **Plaintext SQLite settings / folder index** — contents are non-secret;
  confidentiality relies on the per-user `%LocalAppData%` ACL (the correct OS
  trust boundary). No app-level encryption planned.
- **External-recipient and diagnostics dialogs** that display addresses/paths
  are **intentional UX**, not information-disclosure leaks.
- **Send-guard fails open** — per the project rule that exceptions must never
  escape into Outlook (which would disable the add-in). A security team wanting
  fail-safe behaviour is addressed by the Phase 9 backlog item above.
