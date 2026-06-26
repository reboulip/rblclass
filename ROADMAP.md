# RBLclass roadmap

**Closed history** — Phases 0–5 and every shipped feedback wave
(**v2.0/v2.1/v2.2/v2.4.0.0/v2.4.1.0/v2.5.0.0**) — lives in
[docs/roadmap-history.md](docs/roadmap-history.md). This file holds **where we
are**, the few **open carried-over items**, and the **forward-looking phases**.

Reference docs from the legacy analysis:
[docs/legacy-overview.md](docs/legacy-overview.md),
[docs/legacy-reimplementation-roadmap.md](docs/legacy-reimplementation-roadmap.md),
[docs/parity-and-regression.md](docs/parity-and-regression.md).

## Where we are (2026-06-26)

The rewritten add-in reached full legacy parity and shipped as **v2.0.0.0**,
followed by feedback/refinement waves **v2.1.0.0**, **v2.2.0.0**, **v2.4.0.0**
(classify performance/freeze fix, classify-after-send, attachment disposition),
**v2.4.1.0**, **v2.5.0.0**, **v2.5.1.0** (diagnostic PERF logging), and
**v2.5.2.0** (classify-time freeze fix: deferred post-classify refresh +
mid-batch selection suppression) — all complete and recorded in
[history](docs/roadmap-history.md).

- [x] **Classify-time freeze — FIXED & verified (v2.5.2.0).** The
      selection-event storm during classify is gone: the 2026-06-26 production
      logs show `SelectionChange suppressed (batch in progress)` on every
      multi-item batch and a single deferred `RefreshSelection` (ran ≤56 ms)
      after each classify; normal classify latency is ~80–650 ms total. This is
      the freeze the user cared about. *(Separately observed, not this issue: two
      ~80–90 s stalls on the attachment-save-to-SharePoint path, and a 1× auto-class
      COMException moving into the account-root inbox — tracked as their own
      follow-ups, not the classify freeze.)*

The only remaining product item is the live re-verification of the
external-sender banner strip (below — not reproducible on the dev machine).
Phase 6 (pilot rollout) is in progress; Phases 7–9 are forward-looking.

## Testing & validation strategy (applies to every item)

- **Automated unit tests:** all non-trivial logic lives in `RBLclass.Core`,
  covered by **xUnit + FluentAssertions + NSubstitute**. The folder-search
  matcher and the classify / send-guard decision logic are the highest-value
  surfaces. The Outlook adapter and add-in shell are excluded (they need a live
  Outlook) and validated manually.
- **Manual validation on this dev workstation:** every user-visible increment
  ships an installable add-in via the install kit; load into the dev Outlook
  and validate by hand before it moves on (`/reload-addin`).
- **Target validation:** parity-critical and COM/EDR/bitness-sensitive
  increments are re-checked on the real 32-bit target before pilot. Items that
  are explicitly **not verifiable on the dev machine** (Stormshield/S-MIME mail,
  the external banner) carry a re-verify note.

---

## Open carried-over items

- [ ] **Strip the external-sender banner on reply and on classify** —
      implemented in the v2.1.0.0 scope; **not verified live** (no
      external-banner mail on the dev machine). Re-verify on a workstation that
      receives the company banner before closing.

---

## v2.6.0.0 — UI refresh & classify depth

### A — UI & ribbon
- [x] **A1.** Ribbon UI overhaul: add a Settings icon, remove the "Remove Attachments" and "Index Status" buttons, add an About button (version / author / tech info), and an easter egg — clicking the hidden item in About replaces the Auto-Class button label with a pig icon. [#15]
- [x] **A2.** Increase the prominence of the Classify button in the task pane. [#17]
- [ ] **A3.** Add an auto-expand toggle and a visual row separator to search results. [#13]

### B — Classify & attachments
- [ ] **B1.** Extend the auto-classify conversation-history retention window and expose it as a configurable setting. [#16]
- [ ] **B2.** Add a "use the same save directory" checkbox (checked by default) to the multi-mail attachment disposition UI. [#18]
- [ ] **B3.** Add diagnostics for external-sender banner detection, improve reliability, and add fine-grained configuration; closes the open carried-over verification item. [#19]

### C — Meeting items
- [ ] **C1.** Support classifying MeetingItem objects from the selection, enabled via an opt-in toggle in Settings. [#12]

---

## Phase 6 — Pilot rollout (in progress)

- [x] Documentation: trilingual (EN/DE/FR) self-contained single-file HTML
      install guide (`docs/installation-guide.html`, commit `beee310`). **Still
      needs** its `[contact / channel]` placeholder filled in (all three
      languages) before it is emailed to pilots.
- [x] Packaging: product relabelled **v2.0.0.0**, rebuilt as
      `RBLclass-2.0.0.0.msi` (`installer\Build-Installer.ps1`), validated by
      install/uninstall on the dev machine.
- [ ] **Open `develop` → `main` PR (#2)** — gating merge on re-running the
      [docs/parity-and-regression.md](docs/parity-and-regression.md) matrix on
      the real 32-bit target, then tagging `v2.0.0.0` on `main` and packaging
      with `/make-release`.
- [ ] Deploy to 5–10 pilot users.
- [ ] Telemetry (opt-in): performance metrics, error reports.
- [ ] Feedback loop, bug fixes.

## Phase 7 — General availability

- [ ] Communication plan.
- [ ] Self-service install page (link to the per-user MSI on the internal HTTPS
      share).
- [ ] Support process.
- [ ] Monitoring of error logs across the fleet.

## Phase 8 — Continuous maintenance (ongoing)

- [ ] Regression test plan executed on every Semi-Annual Enterprise channel
      update (typically January and July). Matrix: install, folder indexing,
      folder search, open-folder, classify (single + multi-destination),
      conversation widening, sub-folder creation, attachment removal, send-time
      guards, sent-item triage, settings persistence — mirrors the parity
      checklist in [docs/legacy-overview.md](docs/legacy-overview.md) §9.
- [ ] Watch the Microsoft 365 roadmap / deprecation notices for VSTO and
      classic Outlook.
- [ ] Annual review of the migration plan toward "new Outlook" if/when that
      becomes a corporate direction.

## Phase 9 — Security hardening (audit 2026-06-07)

A full audit found **no critical or high-severity vulnerabilities** (SQL fully
parameterized, no path traversal, no hardcoded secrets, no crypto, no outbound
network calls, COM lifetime managed via `ComRef<T>`). The items below are
**defense-in-depth and enterprise due-diligence** for a corporate security
review — several will be dropped on risk/benefit for a per-user, no-network,
single-machine add-in, but are recorded so the review trail is explicit.

### Evidence & process (highest value for sign-off)
- [ ] **Threat-model / security architecture doc** (`docs/security.md`): trust
      boundaries, data handled (folder names/paths + transient send-time
      metadata; **no mail bodies persisted**), data-at-rest location
      (`%LocalAppData%\RBLclass\`, per-user ACL), "no telemetry / no network by
      default", HKCU-only / no-admin posture.
- [ ] **Authenticode-sign the DLL + MSI** with the internal-PKI Code Signing
      cert (OID 1.3.6.1.5.5.7.3.3), timestamped — promoting today's optional
      signing to the GA default.
- [ ] **Dependency vulnerability scanning** in CI
      (`dotnet list package --vulnerable --include-transitive` + Dependabot over
      `SQLitePCLRaw`, `Microsoft.Data.Sqlite`, `Serilog`). Pin and re-verify on
      every bump — the COM host has no binding redirects
      ([[com-addin-no-binding-redirects]]).
- [ ] **SAST in CI** once the Phase 1 CI pipeline lands (Roslyn security
      analyzers / CodeQL, security rules as build errors).
- [ ] **SBOM** per release artifact.

### Error handling & data leakage (actionable code changes)
- [ ] **Stop showing raw exception text to users.** `ShowError`
      (`RblClassAddIn.cs`) calls `MessageBox.Show(ex.ToString())`. Show a
      friendly message; log the full exception via Serilog.
- [ ] **Replace silent `catch { }` blocks** in `OutlookMailStore.cs` with
      logged catches (notably the attachment-strip swallow — a filed copy could
      silently retain attachments the user believes were removed,
      [[classify-attachments-on-copy]]; surface that failure to the user). *Note
      v2.4 F2 reworks the attachment path — fold this in there.*
- [ ] **Decide & document send-guard failure behavior.** `Application_ItemSend`
      deliberately **fails open** (CLAUDE.md forbids exceptions escaping into
      Outlook). A security team may want the external-recipient guard to **fail
      safe**. Narrow the catch so a guard *evaluation* error still prompts while
      genuine COM faults stay contained; document the stance.
- [ ] **Confirm & document log hygiene:** assert no recipient addresses or
      mail-body text reach the Serilog sink. Capture the redaction policy in
      `docs/security.md`.

### Policy & tamper resistance
- [ ] **Admin-lockable security settings.** The security-relevant toggles
      (`SendExternalWarning`, `InternalDomains`, `ForgottenAttachmentKeywords`)
      live in the user-writable SQLite `Settings` table. Add an optional **HKLM
      policy override** the security team can push by GPO (admin policy wins).
      *Likely the highest-value enterprise-specific item.*

### Input validation (low risk)
- [ ] **Cap free-text inputs:** length-limit the folder-search query and the
      settings text editors; add an **upper** bound to `MaxResults`
      (`Settings.ParseMaxResults` clamps only the lower bound today).
- [ ] **Validate `InternalDomains` format** (reject entries without a
      `domain.tld` shape).

### Minor / best-practice
- [ ] **Use `SqliteConnectionStringBuilder`** instead of `"Data Source=" +
      _dbPath`.
- [ ] **Drop file/DB paths from user-facing dialogs** unless in an explicit
      diagnostics mode.

### Accepted by design (recorded so they are not re-flagged)
- **Plaintext SQLite settings/folder index** — contents are non-secret;
  confidentiality relies on the per-user `%LocalAppData%` ACL (the correct OS
  trust boundary). No app-level encryption planned.
- **External-recipient and diagnostics dialogs** that display addresses/paths
  are **intentional UX**, not information-disclosure leaks.
