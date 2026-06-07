# RBLclass roadmap

Reference docs produced from the legacy code analysis — read these for
the per-feature detail this phase plan schedules:
- [docs/legacy-overview.md](docs/legacy-overview.md) — feature memo of
  the legacy VBA RBLclass.
- [docs/legacy-reimplementation-roadmap.md](docs/legacy-reimplementation-roadmap.md)
  — the sequential, feature-by-feature rebuild plan (Steps 1–10) and the
  approved deviations.

## Where we are (2026-06-07)

The legacy-reimplementation roadmap (Steps 0–10) is **complete and signed
off** — see the parity checklist and regression matrix in
[docs/parity-and-regression.md](docs/parity-and-regression.md). Phases 1–4
below map onto those steps and are done, modulo the recorded deviations
called out in each phase (Quick Open dropped, attachment management folded
into Step 4 rather than shipping as its own slice, incremental re-indexing
landed narrower than originally planned, reminder-to-front dropped). Phase
5 ("smart features") was never required for legacy parity and is deferred
indefinitely as a stretch-goal backlog — re-scope it against real pilot
feedback rather than building it speculatively.

Phase 6 (pilot rollout) is **in progress**: the product has been
relabelled **v2.0.0.0** (the first ship of the rewritten add-in,
succeeding the legacy VBA macro — `develop` was 28 commits ahead of `main`
going into this release), packaged as `RBLclass-2.0.0.0.msi`
(`installer\Build-Installer.ps1`, validated on the dev machine), and
documented in a trilingual (EN/DE/FR) self-contained install guide
(`docs/installation-guide.html`). An open `develop` → `main` PR (#2) is
awaiting the 32-bit-target regression pass before merge, tag, and
`/make-release` packaging.

## Testing & validation strategy (applies to every phase)

- **Automated unit tests:** all non-trivial logic lives in
  `RBLclass.Core` and is covered by **xUnit + FluentAssertions +
  NSubstitute**, run in CI on every push. The folder-search matcher and
  the classify/send-guard decision logic are the highest-value test
  surfaces. The Outlook adapter and add-in shell are excluded from CI
  (they need a live Outlook) and are validated manually instead.
- **Manual validation on this dev workstation:** every phase that
  produces a user-visible increment ships an **installable add-in** via
  the POC PowerShell installer (HKCU registration, no admin, no
  signature required — as validated in Phase 0). `Install-RBLclass.ps1` /
  `Uninstall-RBLclass.ps1` are first-class build outputs so any increment
  can be loaded into the dev Outlook and validated by hand before it
  moves on.
- **Target validation:** parity-critical increments are re-checked on
  the real 32-bit target before pilot.

## Phase 0 — De-risking (1 week)

### Already validated ✅
- Outlook target confirmed: Outlook **32-bit**, Semi-Annual Enterprise
  channel. No additional internal admin restriction on COM add-ins
  (third-party add-in Stormshield is loaded and running). No GPO blocks
  add-ins or macros loading. User has write access to
  HKCU\Software\Microsoft\Office\Outlook\Addins\.
- **"Hello PST" COM Add-in POC built and validated end-to-end on the
  real 32-bit target workstation** (ribbon button, PST enumeration,
  `Microsoft.Data.Sqlite` + native `e_sqlite3.dll` loading at both
  bitnesses via the `SQLitePCLRaw.bundle_e_sqlite3` both-arch payload).
- **Per-user install mechanism validated.** An **unsigned** PowerShell
  installer writing HKCU COM registration installs and loads without
  admin rights under the workstation's existing PowerShell execution
  policy, with no Stormshield/EDR objection and no `LoadBehavior`
  flip from 3 to 2.
- **No code signature is required for the add-in to load.** Authenticode
  signing is therefore treated as optional defence-in-depth (and an
  internal-trust nicety), **not** a functional gate. Internal PKI is
  available if/when we choose to sign.

### Notes on signing & packaging (reconciled with the validated POC)
- ClickOnce is **not used** — it is VSTO-specific and excluded for this
  project (see CLAUDE.md). The packaging path is: **unsigned PowerShell
  installer for the POC → WiX per-user MSI for GA**, with signing
  optional on both.
- Confirming the internal PKI can issue a **Code Signing EKU**
  certificate (OID 1.3.6.1.5.5.7.3.3) is a **nice-to-have** to action
  only if/when we decide to sign — it does not block any build or
  deployment.

### Status
- [x] POC built, deployed, and validated on the 32-bit target.
- [x] Go/No-Go gate passed — proceeding to Phase 1.

## Phase 1 — Foundations (2-3 weeks)

Goal: clean architecture, no business value yet, but everything in place
to build fast afterwards.

### Status — done, bar CI ✅

- [x] Solution skeleton (Step 0, commit `481f0b3`):
    - RBLclass.Core (.NET Standard 2.0) — business logic, no Outlook dep
    - RBLclass.Outlook.Adapter (.NET Framework 4.8) — IMailStore impl over COM
    - RBLclass.AddIn (.NET Framework 4.8) — thin COM-add-in shell
      (IDTExtensibility2 + IRibbonExtensibility)
    - RBLclass.Core.Tests (xUnit, 67 tests passing)
- [x] Core interfaces defined: `IMailStore`, `IFolderTree`,
      `IFolderSearch`, `IClassifier`, `ISettingsStore`, plus
      `IFolderRepository` for the SQLite seam (Step 1 architecture
      decision — [[step1-architecture-decisions]]). `IMailIndex` remains
      **deferred** to Phase 5, which is itself now deferred indefinitely
      (see "Where we are" above).
- [x] SQLite schema v1: `Folders` (StoreId, EntryId, ParentEntryId,
      Name, FullPath, IsLeaf), `Settings`, `SchemaVersion` migrations
      table (Step 1). The mail FTS5 index was correctly **not** built —
      legacy parity never needed it.
- [x] Logging infrastructure (Serilog rolling file sink,
      `%LocalAppData%\RBLclass\logs\`)
- [ ] **CI — still open.** No `.github/workflows` yet; the 67
      `RBLclass.Core` tests run and pass locally on every change but are
      not wired into a pipeline. Lowest-risk remaining Phase 1 item —
      revisit once the v2.0.0.0 ship settles.
- [x] POC PowerShell installer promoted to
      `Install-RblClassAddIn.ps1` / `Uninstall-RblClassAddIn.ps1`
      (`.claude/skills/make-release/`) — the single shared identity
      behind both `/make-release` and the `/reload-addin` dev loop, and
      a first-class build output for every increment.
- [x] COM lifetime management utilities (`ComRef<T> : IDisposable`,
      `src/RBLclass.Outlook.Adapter/ComRef.cs`)
- [x] WiX MSI installer — built out fully (not just a skeleton),
      per-user, registers the add-in under
      `HKCU\Software\Microsoft\Office\Outlook\Addins\`, and **validated
      by install/uninstall on the dev machine** (commit `0b17715`,
      rebuilt as `RBLclass-2.0.0.0.msi` for the relaunch). Per CLAUDE.md
      "Packaging", the PowerShell kit and the MSI coexist until the MSI
      is also validated on the 32-bit target — the one open item before
      the kit can retire.

## Phase 2 — Folder indexing engine (2 weeks)

Legacy parity needs a **folder** index only — the legacy tool never
indexed mail bodies (see [docs/legacy-overview.md](docs/legacy-overview.md)).
This phase delivers Steps 1–2 of the reimplementation roadmap. The mail
content index is deferred to Phase 5 (now itself deferred indefinitely).

### Status — done, with two recorded deviations ✅

- [x] Initial folder-tree walk of all open PST stores on **first run
      only**, on a background thread, with progress + cancellation
      (`FolderIndexService`, Step 1 — replaces the legacy modal "please
      wait" form).
- [x] **Persistent** folder index in SQLite, keyed by (StoreId, EntryId).
      Subsequent starts load from SQLite, not Outlook (Step 1).
- [x] *(landed narrower than planned)* **Targeted re-index on
      user-triggered sub-folder creation** —
      `IFolderTree.ReindexStore` wired from `IMailStore.CreateSubfolder`
      (landed in Step 4, not here). What did **not** land: passive
      Outlook `FolderAdd`/`FolderRemove`/`FolderChange` event
      subscription with a background reconcile. No drift has surfaced in
      dev-machine validation across Steps 1–10; revisit only if the
      32-bit-target regression pass finds the index going stale between
      sessions.
- [x] Configurable store/folder exclusions (`FolderExclusionPolicy`,
      Step 1 — replaces the legacy hard-coded FR/EN substrings).
- [x] *(landed as a configurable option, not a single fixed mode)*
      Folder keyword search (`FolderSearchService`/`IFolderSearch`): AND
      across keywords, case- and accent-insensitive, collapse/expand,
      result cap — matched over the SQLite-hydrated tree, not FTS5, as
      planned. The match mode itself is now a **user setting**: it
      defaults to **word-prefix** (closer to legacy behaviour for the
      common case) with **substring** as an opt-in (Step 9,
      [[folder-search-match-mode]]) — a deliberate refinement over the
      original "always substring" plan, approved during Step 9.
- [x] Performance target (folder search < 100 ms) verified live on the
      dev-machine tree during Steps 2 and 10 sign-off.

## Phase 3 — Core UX (3 weeks)

Delivers reimplementation roadmap Steps 3–6 (open-folder, classify,
quick-open, conversation widening + task guard) and Step 9 (settings).

### Status — done, with one drop and one stack deviation ✅

- [x] Custom Ribbon (Ribbon XML) with main actions
- [x] *(deviation from this plan, approved & documented in CLAUDE.md)*
      WPF Task Pane with a **hand-rolled minimal MVVM**
      (`RBLclass.AddIn/Mvvm`: `ObservableObject` + `RelayCommand`),
      bridged via `ElementHost` into a ComVisible WinForms CTP host —
      **not** CommunityToolkit.Mvvm as originally planned here. Its 8.x
      dependency graph (`System.Memory` ≥ 4.5.5,
      `System.Runtime.CompilerServices.Unsafe` ≥ 6.0.0) would float the
      SQLite facade assemblies above what `SQLitePCLRaw` was built
      against, and a COM host has no binding redirects to fix that at
      runtime (commit `f66e2b1`, [[com-addin-no-binding-redirects]]).
- [x] Fast folder picker over IFolderSearch (open-a-folder, Step 3)
- [x] "Classify selected mail(s)" command — multi-destination,
      copy-per-destination, attachment handling on copy
      ([[classify-attachments-on-copy]]) (Step 4)
- [x] Inline sub-folder creation + targeted re-index (Step 4)
- [x] Conversation widening + task-completion guard (Step 6)
- [x] ~~Quick Open "jump to last filed folder" affordance (Step 5)~~ —
      **DROPPED, not deferred.** A pane-local infobar variant was built
      and tried live; it didn't land well and is **not to be
      re-attempted** as originally specified
      ([[step5-quick-open-dropped]]).
- [x] Settings pane backed by the SQLite `Settings` table — replaces the
      legacy positional `options.txt` (Step 9: typed `Settings` in Core +
      a WPF Options-style `SettingsWindow`; eleven options total — the
      legacy nine minus Quick Open, plus `FolderMatchMode`,
      `InternalDomains`, `ForgottenAttachmentKeywords` —
      [[settings-always-persistent]])

## Phase 4 — Productivity features (2 weeks)

Delivers reimplementation roadmap Steps 7, 8 (and folds in attachment
management, which landed inside Step 4 rather than as its own slice — see
below).

### Status — core deliverables done; one item dropped from scope ✅

- [x] *(landed inside Step 4, not as a separate Step 5 slice)*
      Attachment management — standalone "Remove attachments" command +
      the classify-time toggle, with **filed-copy-only** semantics: the
      "remove attachments" action strips the filed copy only, never the
      original; "keep a copy" keeps the original's attachments intact
      ([[classify-attachments-on-copy]]).
- [x] Configurable `ItemSend` guards (Step 7, commit `e87a386`):
      `ForgottenAttachmentGuard` (configurable keyword list) and
      `ExternalRecipientGuard` (`AddressEntry` type + configurable
      internal-domain allowlist — **not** the legacy `CORPMAIL`
      substring), both backed by Settings.
- [x] Sent-item triage prompt — Class / Delete / Move-to-Inbox / Leave,
      whole-conversation, with a clean re-entrancy guard
      (`SentItemTriageWindow`/`ViewModel`, Step 8, commit `08226dc`)
- [ ] ~~Keyboard shortcuts for top folders (most used / pinned)~~ — **not
      built**, and not part of the legacy parity checklist. Leave it out
      of the initial release; only revisit if pilot feedback asks for it
      (avoid building speculative UI ahead of real usage signal).

## Phase 5 — Mail content index + smart features (2 weeks)

> **Status: deferred indefinitely**, not just postponed. Steps 1–10 of
> the reimplementation roadmap reached full legacy parity without any of
> this — the legacy tool never indexed mail bodies (deviation #1 in
> [docs/legacy-reimplementation-roadmap.md](docs/legacy-reimplementation-roadmap.md)).
> Treat this phase as a stretch-goal backlog to be **re-scoped against
> real pilot feedback** once v2.0.0.0 ships, rather than built ahead of
> demand.

- [ ] **Deferred mail index** (IMailIndex): SQLite **FTS5** over mail
      subject/body, keyed by (StoreId, EntryId), with incremental updates
      via mail **item** events (ItemAdd / ItemChange / ItemRemove). This
      was pulled out of the parity track (legacy never indexed mail); it
      lands here because the smart features below need it.
      Performance target: search < 100 ms on 100k mails.
- [ ] Conversation-based auto-classification suggestion (local rules,
      SQLite-backed)
- [ ] IClassificationSuggester abstraction
- [ ] Optional Azure OpenAI implementation behind a feature flag
- [ ] Response caching in SQLite

## Phase 6 — Pilot rollout (2 weeks)

### Status — packaging & docs done; deployment pending target validation 🚧

- [x] Documentation: trilingual (EN/DE/FR), self-contained single-file
      HTML install guide for end users
      (`docs/installation-guide.html`, commit `beee310`) — covers what
      pilots receive, the MSI install flow (no admin rights, SmartScreen
      guidance), uninstall, log location, and a caveats table. **Still
      needs** its `[contact / channel]` placeholder filled in (all three
      languages) before it is actually emailed to pilots.
- [x] Packaging: product relabelled **v2.0.0.0** (the first ship of the
      rewritten add-in — see "Where we are" above) and rebuilt as
      `RBLclass-2.0.0.0.msi` via `installer\Build-Installer.ps1`,
      validated by install/uninstall on the dev machine.
- [ ] **Open `develop` → `main` PR (#2)** — gating merge on: re-running
      the [docs/parity-and-regression.md](docs/parity-and-regression.md)
      regression matrix on the real 32-bit target workstation, then
      tagging `v2.0.0.0` on `main` and packaging the validated build with
      `/make-release`.
- [ ] Deploy to 5-10 pilot users
- [ ] Telemetry (opt-in): performance metrics, error reports
- [ ] Feedback loop, bug fixes

## Phase 7 — General availability

- [ ] Communication plan
- [ ] Self-service install page (link to the per-user MSI on the
      internal HTTPS share)
- [ ] Support process
- [ ] Monitoring of error logs across the fleet

## Phase 8 — Continuous maintenance (ongoing)

- [ ] Regression test plan executed on every Semi-Annual Enterprise
      channel update (typically January and July). Test matrix: install,
      folder indexing, folder search, open-folder, classify
      (single + multi-destination), conversation widening, sub-folder
      creation, attachment removal, send-time guards (forgotten
      attachment + external recipient), sent-item triage, settings
      persistence. Mirrors the parity checklist in
      [docs/legacy-overview.md](docs/legacy-overview.md) §9.
- [ ] Watch Microsoft 365 roadmap and deprecation notices for VSTO and
      classic Outlook.
- [ ] Annual review of the migration plan toward "new Outlook" if/when
      that becomes a corporate direction.
