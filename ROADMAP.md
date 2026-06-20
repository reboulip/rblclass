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

**Update (2026-06-09):** the throwaway Phase 0 POC has been removed now that
the product is stable, and the first wave of operational pilot feedback is
scoped for **v2.1.0.0** — see the planned section at the end of this file.

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

## Phase 9 — Security hardening (audit 2026-06-07)

A full audit of `src/` and `tests/` found **no critical or high-severity
vulnerabilities**: SQL is fully parameterized (`SqliteFolderRepository`,
`SqliteSettingsStore` — every query uses `$param` binding, no string-built
SQL), there is no filesystem path traversal (DB/log paths come from
`Environment.GetFolderPath(LocalApplicationData)` + `Path.Combine`; folder
"paths" are Outlook OM data, not filesystem), folder search is in-memory over
the cached tree (`FolderSearchService` — no query-time SQL/FTS injection
surface), there are no hardcoded secrets, no cryptography to weaken, and no
outbound network calls (Azure OpenAI stays Phase-5 deferred). COM object
lifetime is managed via `ComRef<T>` and events are unsubscribed on shutdown.

The items below are therefore **defense-in-depth and enterprise
due-diligence** to clear a large-enterprise security review — not fixes for
known exploits. The stated bar is approval by a corporate cybersecurity team;
several items will be **dropped on risk/benefit** for a per-user, no-network,
single-machine add-in, but they are recorded so the review trail is explicit.
For enterprise sign-off the *evidence* items (threat-model doc, signing,
dependency scanning) are usually worth more than the code tweaks.

### Evidence & process (highest value for enterprise sign-off)

- [ ] **Threat-model / security architecture doc** (`docs/security.md`):
      trust boundaries, data handled (folder names/paths + transient mail
      metadata at send time; **no mail bodies are persisted** — the FTS5 mail
      index was deliberately never built), data-at-rest location
      (`%LocalAppData%\RBLclass\`, per-user ACL), an explicit "no telemetry /
      no network calls by default" statement, and the HKCU-only / no-admin
      install posture. This is usually the single artifact a security team
      asks for first.
- [ ] **Authenticode-sign the DLL + MSI** with the internal-PKI Code Signing
      cert (OID 1.3.6.1.5.5.7.3.3), timestamped — promoting today's *optional*
      signing (Phase 0 notes) to the GA default. Proves integrity/provenance
      and removes SmartScreen friction. (Signing is not a functional gate; this
      is a trust nicety the security team will expect anyway.)
- [ ] **Dependency vulnerability scanning** in CI:
      `dotnet list package --vulnerable --include-transitive` plus GitHub
      Dependabot / advisory alerts over `SQLitePCLRaw`, `Microsoft.Data.Sqlite`
      and `Serilog`. Pin and re-verify versions on every bump — the COM host
      has no binding redirects, so transitive strong-named versions must match
      exactly ([[com-addin-no-binding-redirects]]).
- [ ] **SAST in CI** once the still-open Phase 1 CI pipeline lands: enable
      Roslyn security analyzers / CodeQL on every push and treat security rules
      as build errors.
- [ ] **SBOM** generated for each release artifact (supply-chain inventory the
      security team can ingest).

### Error handling & data leakage (actionable code changes)

- [ ] **Stop showing raw exception text to users.** `ShowError`
      (`RblClassAddIn.cs:762-768`) calls `MessageBox.Show(ex.ToString())`,
      exposing stack traces and internal file paths on screen. Show a friendly
      message and log the full exception via Serilog (`TryLog`) instead.
- [ ] **Replace silent `catch { }` blocks** in `OutlookMailStore.cs` with
      logged catches (e.g. attachment removal/save at lines 504 & 508, the
      generic swallow at 562, folder ops at 209/216, store enumeration at 49).
      The attachment-strip swallow is the notable one: a filed copy could
      silently retain attachments the user believes were removed
      ([[classify-attachments-on-copy]]) — log every failure and surface the
      attachment-removal failure to the user.
- [ ] **Decide & document send-guard failure behavior.** `Application_ItemSend`
      (`RblClassAddIn.cs:425-444`) deliberately **fails open** — its top-level
      try/catch never blocks a send, because CLAUDE.md forbids letting an
      exception escape into Outlook (it would flip `LoadBehavior` 3→2). A
      security team may want the external-recipient guard to **fail safe**
      (warn/block) instead. Resolve the tension explicitly: narrow the catch so
      a guard *evaluation* error still prompts the user, while genuine COM
      faults stay contained. Document whichever stance is chosen.
- [ ] **Confirm & document log hygiene:** assert that no recipient addresses or
      mail-body text reach the Serilog sink (`RblClassAddIn.cs:745-755`,
      Information level, 14-day retention). Capture the redaction policy in
      `docs/security.md`.

### Policy & tamper resistance (enterprise control)

- [ ] **Admin-lockable security settings.** The security-relevant toggles
      (`SendExternalWarning`, `InternalDomains`, `ForgottenAttachmentKeywords`)
      live in the user-writable SQLite `Settings` table, so a user — or malware
      running at user scope — can disable the external-recipient warning. Add an
      optional **HKLM policy override** the security team can push by GPO to
      enforce the external guard fleet-wide (admin policy wins over the user
      setting). *Likely the highest-value enterprise-specific item.*

### Input validation (low risk — single-user, 32-bit memory budget)

- [ ] **Cap free-text inputs:** length-limit the folder-search query
      (`FolderSearchService.Tokenize`) and the settings text editors
      (`InternalDomains`, `ForgottenAttachmentKeywords`), and add an **upper**
      bound to `MaxResults` (`Settings.ParseMaxResults`, `Settings.cs:74-80`
      currently clamps only the lower bound at 1). Guards against
      self-inflicted memory pressure inside the ~1 GB 32-bit Outlook budget.
- [ ] **Validate `InternalDomains` format** (reject entries without a
      `domain.tld` shape) so a malformed allowlist can't silently weaken the
      external-recipient guard.

### Minor / best-practice

- [ ] **Use `SqliteConnectionStringBuilder`** instead of
      `"Data Source=" + _dbPath` (`RblClassAddIn.cs:147`). The path is fixed and
      not user-influenced today, so this is purely hygiene.
- [ ] **Drop file/DB paths from user-facing dialogs** unless in an explicit
      diagnostics mode — the folder-index status dialog currently prints
      `_dbPath` and `_logDirectory`.

### Accepted by design (recorded so they are not re-flagged)

- **Plaintext SQLite settings/folder index** — contents are non-secret
  (toggles, folder names/paths); confidentiality relies on the per-user
  `%LocalAppData%` ACL, which is the correct OS trust boundary. No app-level
  encryption is planned.
- **External-recipient and diagnostics dialogs** that display addresses/paths
  are **intentional UX** (the whole point of the external-recipient guard),
  not information-disclosure leaks.

## v2.1.0.0 — Pilot-feedback UX & fixes (planned)

The first feedback wave from operational use on the test machine (Phase 6
"Feedback loop, bug fixes"). All of it is polish, UX, and bug-fix work on
top of the shipped v2.0.0.0 — no new architectural layer. Ordered so the
**theming/layout foundations** land first (they touch the views everything
else styles), then the **isolated correctness/UX fixes**, then the three
**design-heavy redesigns** last (each carries open implementation questions
to resolve when it's picked up). Per CLAUDE.md, any non-trivial `RBLclass.Core`
logic added here (search default, classify-to-root, banner signature matching
if it lands in Core) ships with xUnit coverage.

Decisions already taken with the maintainer are baked into each item below;
the remaining **open questions** are called out inline.

### A. Theming & layout foundations

- [x] **Theme-aware task-pane & dialog styling (fixes the readability +
      dark-theme reports).** Today some CTP text is dark-on-dark-grey and the
      search-results surface stays white under a dark Outlook theme. Read the
      **Office UI theme** (White / Colorful / Dark Gray / Black) and apply a
      matching background+foreground palette across the task pane **and every
      modal window** (settings, folder picker, send guards, sent-item triage).
      **Caveat:** the Office theme can itself be set to *"Use system
      setting"* — in that case resolve the actual light/dark palette from the
      **Windows app theme**. Centralise the colours in a shared WPF
      `ResourceDictionary` (theme brushes) consumed by all views so nothing
      hard-codes white/black again.
      - Open Qs: react to a live theme switch while Outlook is running, or
        resolve once at pane/dialog creation? Exact registry/API source for
        the Office theme vs the Windows theme.
- [x] **Long folder-path display in the narrow vertical pane.** Folder paths
      can exceed the CTP width; the informative end of the path (the leaf and
      its immediate parents) must stay visible. Left-truncate with a leading
      ellipsis (show the *end*, not the start) and expose the full path on
      hover/tooltip. Applies to the unified list (below) and any path label.

### B. Isolated correctness & UX fixes

- [x] **Make "contains" the default folder-search match mode.** Search
      currently matches a keyword as a word *prefix*, so `security` misses
      `Cybersecurity`. Flip the default `FolderMatchMode` to substring/contains
      (the mode already exists, `FolderSearchService` / [[folder-search-match-mode]]);
      word-prefix becomes the opt-in. Update the default in `Settings`, applied
      to both the Open and Classify searches. xUnit: `security` matches
      `Cybersecurity`; word-prefix still available via the setting.
- [x] **Allow classifying into a PST store's root node.** Users want to file
      directly at the top of a `.pst`, above any folder. Make the **store root
      node** a selectable destination in the folder tree / search results /
      unified list, and support `Move` into the store root in the adapter.
      - Open Q: do store roots also participate in keyword search results, or
        only appear as an always-available "(root of <store>)" entry?
- [x] **Freeze the pane while a classify/move is running (anti-double-fire).**
      A slow classify tempts a second Enter / Classify click, which today
      queues an action that then fires against the *next* mail selected after
      processing. While a classify or move is in flight, disable pane input
      (search box, list, toggles, buttons, Enter) and show a busy indicator;
      ignore repeat triggers; re-enable on completion. A clean re-entrancy
      guard in `ClassifyViewModel.DoClassify`.
- [x] **Process encrypted conversation siblings correctly (Stormshield).**
      *(Implemented; NOT verified live — no encrypted test message available on
      the dev machine. Covered by unit tests only; re-verify on a workstation
      that has S/MIME / Stormshield mail before relying on it for pilot.)*
      When "process same conversation" is on and Stormshield is **not active**,
      encrypted siblings of the conversation are currently dropped from the
      run. Decision: **skip them but surface a notice** listing the items that
      were not processed (rather than silently omitting them) — do not fail the
      whole classify. Wired into the conversation-widening path used by classify
      and sent-item triage: `GetConversationSiblings` now returns a
      `ConversationSiblings` (processable + skipped-encrypted subjects), carried
      through `ClassifyPreflight` to a pane status note and a triage message box.
- [x] **Position modal windows on the active Outlook monitor (multi-screen).**
      *(Implemented; NOT verified live — single-monitor dev machine. Re-verify on
      a multi-monitor workstation.)* The "process sent item" modal (and the other
      app modals) can open on the primary monitor while Outlook lives on another.
      Each WPF dialog (settings, folder picker, sent-item triage) is now owned by
      the Outlook main window (`DialogPlacement.OwnByOutlook` via the process
      main-window handle) and uses `WindowStartupLocation="CenterOwner"`, so it
      centres over Outlook wherever it lives, falling back to the primary screen
      if the handle can't be resolved. Plain Win32 message boxes are out of scope.
- [x] **Pick up folders created/renamed directly in Outlook, via a manual
      Refresh.** Reported in pilot testing: creating (or renaming) a sub-folder
      by hand in the Outlook tree does not surface it in folder search — only
      the add-in's own "New subfolder" action re-indexes. **Decision: add a
      "Refresh folders" ribbon button** that re-walks the live stores on demand
      (`IFolderTree.WalkAndPersist`, on the Outlook UI thread — the same path as
      the first-run walk) and refreshes the cache, so created/renamed/deleted
      folders all reconcile in one go. *Rejected for now,* with reasons recorded
      so the trade-off is explicit: **live folder events** — Outlook has no
      tree-wide folder-change event, so catching nested creation/rename needs
      recursive `FolderAdd`/`FolderChange`/`FolderRemove` sinks on every folder
      held all session: too much COM-lifetime sprawl and memory for the 32-bit
      process, for an infrequent action; **periodic timer walk** — wasteful full
      walks regardless of change, still stale between ticks. The button is
      on-demand, robust, and adds no standing COM overhead. Revisit auto-sync
      only if manual refresh proves painful. Underlying walk is already covered
      by `FolderIndexServiceTests`; the button itself is shell wiring.

### C. Design-heavy redesigns (open questions inline)

- [x] **Unify the two pane modes into one always-open task pane.** Today a
      single CTP swaps its whole body between an Open view and a Classify view
      and the ribbon button toggles it shut — so it *feels* like two panes that
      keep vanishing. Replace both with **one always-open pane built around a
      single folder-search list**, with these row interactions:
      - **single-click** → toggle the folder into the classify selection (check);
      - **double-click** → file to that one folder immediately;
      - **per-row button** → open/reveal that folder in the Outlook explorer.

      The top of the pane carries the selected-mail count, the option toggles
      (keep-a-copy, remove-attachments, widen-conversation, all-results), and a
      **"Classify to N folders"** button (and Enter) for filing into all checked
      folders at once. The ribbon "Open folder" / "Classify" buttons reveal/focus
      the pane instead of toggling it closed. Retire the `PaneMode` /
      `ShowOpenFolder` / `ShowClassify` split (`RblClassTaskPaneHost`,
      `TaskPaneServices`, `RblClassAddIn.TogglePane`) in favour of the merged
      view ([[prefer-isolated-new-ui-over-retrofit]] — build the merged list as a
      new view rather than retrofitting either existing one).
      - Resolved: a **single "RBLclass pane" ribbon toggle** (the two-button
        split retired); **per-row "+"** for sub-folder creation (prompts via a
        small `NamePromptWindow`); **Enter files to the highlighted/first** result
        (double-click files to the row under the cursor). Shipped as new
        `MainPaneView` + `MainPaneViewModel` + `NamePromptWindow`; old
        `FolderSearchView`/`ClassifyView` (+ their VMs) and `PaneMode` deleted.
- [x] **Rework the sent-item triage into a dropdown setting.** Replace the
      on/off `SentItemTriagePrompt` with a **dropdown**: *Move-to-Inbox /
      Delete / Leave / "Let me choose"*. A fixed value applies **automatically
      with no modal**; *"Let me choose"* shows the modal with just those three
      buttons. **Remove the "Class" action and the whole-conversation
      checkbox** from triage (they no longer fit this model). Includes a
      settings migration from the old boolean key
      (`SentItemTriageViewModel` / `SentItemTriageWindow` / `SettingsKeys`).
- [x] *(implemented 2026-06-13 in the v2.2.0.0 sprint; Core matcher
      unit-tested, but the Outlook wiring — learn-from-selection, the
      reply/forward NewInspector strip, and the classify-time strip — is
      NOT verifiable on this machine, which has no external-banner mail.
      Re-verify live on a workstation that receives the company banner.)*
      **Strip the external-sender banner on reply and on classify.** The ugly
      "external email" reminder banner varies per company, so RBLclass learns
      it from a **sample the user teaches it once** (capture a real banner —
      text/HTML — derive a signature, strip matching blocks). Two triggers:
      1. **On reply / reply-all / forward** — auto-strip the banner from the
         draft when a settings toggle is on.
      2. **At classify time** — a tickbox in settings (default: checked) determines
         whether the banner should be stripped from the **filed copy**.
      - Open Qs (detailed design needed): the capture UX (paste vs "pick from
        an open mail"); how robust the signature is (HTML structure vs plain
        text, localised/variant banners, more than one banner per company);
        filed-copy-only vs also-original semantics
        ([[classify-attachments-on-copy]] sets the precedent: act on the copy,
        leave the original intact); whether banner matching belongs in
        `RBLclass.Core` (testable) or the Outlook adapter (HTML body access).

## v2.2.0.0 — Second feedback wave (sprint started 2026-06-12)

Eight items from continued operational use, analysed and sequenced
2026-06-12. Ordering: quick search-correctness fixes first, then the
classify-engine rework (the Stormshield fix) that Undo and Auto-class
build on, then the feature work. Every task ends with a build,
`/reload-addin`, and a manual verification pass in Outlook before the
next one starts. Non-trivial `RBLclass.Core` logic ships with xUnit
coverage per CLAUDE.md. Product version bumps to 2.2.0.0 at ship.

Decisions taken with the maintainer on 2026-06-12 are baked into the
items below (multi-search chip UX, Undo in the pane, Auto-class on the
ribbon, and the encrypted-attachment rule).

### A. Search correctness & responsiveness

- [x] *(verified live 2026-06-12)* **Fix special characters in folder
      search ("R&D").** Reported: a
      folder named `R&D` is not found by the query `R&D`. Root cause: in
      word-prefix mode `FolderSearchService.SplitWords` keeps only
      letters/digits, so the path word list for "R&D" is `r`,`d` and the
      token `r&d` can never prefix any word. Fix: a token containing a
      non-alphanumeric character falls back to substring containment (by
      definition it cannot be a word prefix). xUnit: `R&D` matches in
      both match modes; plain tokens keep word-prefix semantics.
- [x] *(verified live 2026-06-12)* **Minimum query length + typing
      debounce.** Today the first
      keystroke already searches, matching a huge folder set and
      stuttering the pane. Two new persisted settings (Settings dialog,
      "Folder search" group; [[settings-always-persistent]]):
      - `MinSearchLength` — default **2**, search runs only once the
        query reaches it (enforced in `FolderSearchOptions`/Service so
        it is unit-testable).
      - `SearchDebounceMs` — default **200 ms** (clamped 0–2000):
        re-search fires only after the user stops typing for that long
        (`DispatcherTimer` in `MainPaneViewModel`; checkbox toggles and
        sub-folder creation still refresh immediately).
- [x] *(verified live 2026-06-12)* **Ctrl toggles "List every matching
      folder" from the search box.**
      With focus in the query box, pressing and releasing Ctrl *alone*
      toggles the AllResults checkbox; Ctrl+key combinations (Ctrl+A,
      Ctrl+C…) are not affected.
- [x] *(verified live 2026-06-12)* **"Open folders in a new window"
      becomes settings-only.** The
      checkbox leaves the task pane (it is configuration, not a
      per-action choice); the Settings dialog remains the single editor
      and the pane reads the stored value live at navigation time.
- [x] *(verified live 2026-06-12)* **Select all text when the search box
      gains focus by click.**
      (Added 2026-06-12.) When focus is *outside* the search bar and the
      user clicks into it, the existing query is fully selected so
      typing replaces it — sequential classifying nearly always starts a
      fresh keyword search, and manually erasing the old query every
      time is friction. Clicking again once focused keeps normal caret
      behaviour (so the user can still edit/position within the text).

### B. Classify engine rework

- [x] *(implemented 2026-06-13; move/copy/strip/triage paths verified live
      on the dev machine. NOT yet verifiable here: the Stormshield error
      itself and the encrypted-skip rule - no Stormshield/S-MIME mail on
      this machine; re-verify both on the 32-bit target before pilot.)*
      **Stop provoking Stormshield's `MAPI_E_NOT_FOUND` on classify.**
      Reported error (`IMessage.GetAttachmentTable: MAPI_E_NOT_FOUND` in
      `Arkoon.SecurityBox...OnBeforeReadAsync`) is thrown by the
      Stormshield add-in's own async reader against items our classify
      creates and immediately destroys: `mail.Copy()` materialises a
      transient copy in the *source* folder which `Move` then yanks
      away, and the original is afterwards `Delete`d into Deleted Items
      — by the time Stormshield's `OnBeforeRead` inspects either, the
      underlying MAPI message is gone. Fix — make classify move-based:
      - KeepCopy **off**: copies are made only for destinations 1…n−1;
        the original is **moved** to the last destination (single
        destination ⇒ a pure `Move`, no transient copy, no delete, no
        Deleted Items churn). `IMailStore` gains `MoveItemToFolder`.
      - KeepCopy **on**: unchanged (copy per destination, original
        untouched).
      - Sent-item triage "Move to Inbox" rides the same path and becomes
        a true move.
      - **Decision (2026-06-12): encrypted mail (`IPM.Note.SMIME`) is
        never stripped of attachments** — its attachments *are* the
        message. Applies to classify-time stripping and the standalone
        "Remove attachments" command alike; skipped items are reported.
        Accepted trade-off: for regular mail with KeepCopy off, "remove
        attachments" now strips the moved item itself (nothing lands in
        Deleted Items any more — the user accepted the original's
        deletion in that configuration anyway).
      - xUnit: ClassifierService move/copy orchestration for 1 and n
        destinations × KeepCopy on/off; encrypted-skip behaviour.
      - **Decision (2026-06-13, after live verification):** the old "a
        copy always lands in Deleted Items" side effect returns as an
        **opt-in setting** (`ClassifySafetyCopy`, default off): when on
        (and keep-a-copy off), each moved original also leaves a copy in
        its source store's Deleted Items - taken from the moved item *at
        its destination* (never a transient in the displayed folder, so
        the Stormshield race is not re-created) and before attachment
        stripping. Best-effort: its failure never fails the filing.
        Undo (below) is the designed guardrail; this is for users who
        relied on the Deleted Items copy.

### C. Filing features

- [x] *(verified live 2026-06-13)* **Accumulate destinations across
      successive searches.** Checked
      folders persist in a selection set (keyed StoreId+EntryId) that
      survives re-searches: search kw1 → check folders → search kw2 →
      check more → one "Classify to N folders". A chip strip above the
      Classify button lists the selection (per-chip ✕, plus clear-all);
      re-searching re-checks rows already selected; the set clears
      automatically after a successful classify.
- [x] *(verified live 2026-06-13: move-back, copy removal, keep-a-copy
      and safety-copy cleanup all confirmed)* **Undo the last filing
      action.** `ClassifierService.Classify`
      returns an undo plan alongside the result: every move performed
      (with the item's source folder), every copy created, every flag
      set, and (once Auto-class lands) the history rows written. A
      single-slot Undo button next to Classify — enabled only when a
      plan exists — reverses it: moves items back to their source
      folders, deletes the created copies, restores flags, rolls back
      history rows. Stripped attachments cannot be restored; Undo says
      so when that applies. xUnit over a fake store.
- [x] *(verified live 2026-06-13: filed / no-history / undo cases all
      confirmed in the pane)* **Auto-class.** Schema **v2** adds
      `ClassificationHistory(ConversationKey, DestStoreId, DestEntryId,
      WhenUtc)`, appended on every successful classify (conversation key
      = the Outlook `ConversationID`, read before the move invalidates
      the item). Auto-class files each selected mail to its
      conversation's most recent recorded destination(s) — validated
      against the current folder index — honouring the same
      keep-a-copy / remove-attachments / safety-copy settings, and Undo
      covers it like any other classify.
      - **Reworked after the first live test (2026-06-13):** moved off
        the ribbon into a **small "Auto-class" button in the task pane**
        (top-right, by the selection count) and **dropped the modal
        summary** — instead the filed **destination folders are shown in
        the pane's results list** (so the user sees where mail went and
        can open them), and the no-history / stale-folder / error
        outcome is reported in the **pane status line**. The ribbon
        button and its message box were removed.

### Carried over from v2.1.0.0

- [ ] **Strip the external-sender banner on reply and on classify** —
      unchanged scope, still parked behind the items above (open design
      questions recorded in the v2.1.0.0 section).

## v2.4.0.0 — Third feedback wave (sprint started 2026-06-20)

Six items from continued operational use: three quick independent fixes
(A), one pane-interaction redesign (B), one visual redesign (C), and one
bug investigation (D). Sequenced so lower-risk items land first and build
confidence before touching the pane layout. Every task ends with a build,
`/reload-addin`, and a manual verification pass before the next one
starts. Non-trivial `RBLclass.Core` logic ships with xUnit coverage per
CLAUDE.md.

### A. Quick wins

- [x] **A1. Clear the search field after a successful classify.**
      The query persists in the box after each classify, so the user
      has to manually erase it before the next session. After a
      successful classify, set `Query = ""` *silently* — without
      triggering `ScheduleRefresh()` — by writing to the backing field
      and raising `PropertyChanged` directly. The displayed results
      remain visible (the user can still see where the mail landed);
      they clear naturally on the next keystroke once the user types a
      new query. Implementation: add a private `ClearQuerySilently()`
      helper in `MainPaneViewModel`, call it from `DoClassify()` on
      success.
      - **Also drop the select-all-on-focus behavior** from v2.2.0.0
        A5: remove `QueryBox_PreviewMouseLeftButtonDown()` and
        `QueryBox_GotKeyboardFocus()` from `MainPaneView.xaml.cs` (and
        any associated `ClickHandled` tracking state). Because classify
        now clears the query, auto-selecting an already-empty field on
        focus is pointless and can confuse click-to-position use. No
        new settings. No xUnit needed (view-layer behavior only).

- [x] **A2. Async folder index refresh with pane color indicator.**
      The "Refresh folders" ribbon button currently blocks the UI
      thread with a wait cursor and ends with a MessageBox showing
      store/folder counts. Replace with:
      - **Remove the end-of-refresh MessageBox** (`RblClassAddIn.cs`,
        `OnRefreshFoldersClick`). The indicator (below) serves as
        completion feedback.
      - **Non-blocking walk**: `WalkAndPersist` is called on the
        Outlook STA thread (COM objects require it), but `await`-based
        yielding between stores keeps the pane responsive. SQLite
        writes are dispatched to the thread pool.
      - **Color indicator** — a small colored dot added to the pane
        header row (next to the "RBLclass" title or the search box):
        - **Red** — index absent or empty (0 folders; never built,
          or DB cleared). Set at startup before the first-run walk.
        - **Yellow** — walk in progress.
        - **Green** — index populated and ready. Set after
          `WalkAndPersist` completes (both the manual-refresh path
          and the startup first-run walk).
      - `IFolderIndexService` gains an observable `IndexStatus`
        property (`NotFound` / `Indexing` / `Ready`). `FolderIndexService`
        updates it before and after each walk. `MainPaneViewModel`
        subscribes via a property-changed handler and exposes a
        bindable `IndexStatus` property for the indicator.
      - Localization: indicator tooltips (`IndexStatus_NotFound`,
        `IndexStatus_Indexing`, `IndexStatus_Ready`) in
        `Strings.resx` / `.fr.resx` / `.de.resx`.

- [x] **A3. Pane dock position persistence.**
      The task pane always opens docked to the right (`EnsureTaskPane`,
      `msoCTPDockPositionRight` hardcoded). If the user moves it to the
      left, the preference is lost on restart. Fix:
      - Subscribe to `_ctp.DockPositionStateChange` in `RblClassAddIn`;
        on each change, persist the new `DockPosition` integer to a new
        Settings key `PaneDockPosition` (default: `msoCTPDockPositionRight`
        = 2).
      - In `EnsureTaskPane()`, read `PaneDockPosition` from settings and
        set `_ctp.DockPosition` before making the pane visible.
      - Shell wiring only; no xUnit. No localization (no new UI strings).
      - **Pane width is not persisted**: no `WidthChanged` CTP event
        exists, so width detection would require sampling at shutdown —
        more complexity than the benefit warrants for now.

### B. Pane interaction redesign

- [x] **B1. Collapse all checkboxes behind an "Options" toggle.**
      All five per-action checkboxes (AllResults, KeepCopy,
      RemoveAttachments, WidenConversation, StripBanner) are always
      visible but rarely changed per-action; they add visual noise
      below the results list. Replace with a collapsible options panel:
      - **Default state**: the checkboxes StackPanel is hidden
        (`Visibility.Collapsed`); a small, low-profile "Options" button
        (or disclosure triangle `▶ Options`) sits below the results list,
        above the classify button row.
      - **Expand triggers**: (1) user clicks the Options button, or
        (2) user presses Tab while the search box is focused —
        intercept Tab via `PreviewKeyDown` on the search box,
        set `e.Handled = true`, set `IsOptionsExpanded = true`,
        and move keyboard focus to the first checkbox.
      - **Collapse triggers**: clicking Options again (toggle); or a
        successful classify.
      - **On successful classify**: the panel collapses and all checkboxes
        reset to their persisted settings defaults (`KeepCopy`,
        `RemoveAttachments`, `WidenConversation`, `AllResults`, and
        `StripBanner` from their respective `Settings` values). The
        pane is then ready for the next classify session with a clean
        default state.
      - `MainPaneViewModel` gains `IsOptionsExpanded` bool property;
        the XAML panel wraps the checkboxes StackPanel with
        `Visibility="{Binding IsOptionsExpanded, Converter=…}"`.
      - Localization: the Options button label (`Pane_Options`) in
        `Strings.resx` / `.fr.resx` / `.de.resx`.

### C. Search results display

- [x] **C1. Per-row expandable folder path (replaced the hierarchical-display setting).**
      Results currently show the full path as a single left-trimmed
      string (`Personal Archive\Projects\2024\RBLclass`). For deeply
      nested folders the leaf is often cut. Add a second display mode
      where each path segment appears on its own line with increasing
      indentation:
      ```
      Personal Archive
        Projects
          2024
            RBLclass    [↗] [+]
      ```
      - New setting `FolderPathDisplay` (Inline / Hierarchical),
        default: Inline — no change to existing behavior.
      - **Inline mode**: `PathTrim` attached behavior unchanged.
      - **Hierarchical mode**: split `FullPath` on `\` into
        `PathSegments` (string array exposed on `FolderSearchResult`
        or `SelectableFolder`). Render as a `StackPanel` inside the
        `DataTemplate`, with each segment in a `TextBlock` at
        `Margin.Left` increasing by 12 px per depth. The `↗` and `+`
        buttons appear only on the leaf row. The `CheckBox` wraps the
        entire segment stack so any click selects the folder.
        `PathTrim` is bypassed (each individual segment is short).
      - A `DataTemplateSelector` on the results `ListBox` switches
        templates based on `MainPaneViewModel.IsHierarchicalDisplay`.
      - Settings window: new `FolderPathDisplay` toggle/radio in the
        results group, in EN/FR/DE.
      - xUnit: `PathSegments` splitting — root-level folder (no `\`),
        single-level, multi-level, path ending with `\`.

### D. Bug investigation

- [ ] **D1. MAPI_E_NOT_FOUND on multi-item classify (Stormshield).**
      Switching to a pure-move classify in v2.2.0.0 did not eliminate
      the `IMessage.GetAttachmentTable: MAPI_E_NOT_FOUND` error emitted
      by Stormshield's `Arkoon.SecurityBox…OnBeforeReadAsync`. The error
      is **cosmetic** — classify succeeds and mails land correctly — but
      it produces a confusing/worrying notification from Stormshield.
      It only manifests when **more than one email is selected**.

      **Working hypothesis:** Stormshield hooks Outlook item events and
      runs an async scan (`OnBeforeReadAsync`) for each item touched.
      With a single item, our `Move()` call returns and the STA is idle
      before Stormshield's scan runs; the moved item is fully committed
      to the destination store and `GetAttachmentTable` succeeds. With
      multiple items, COM message pumping during the second `Move()`
      call allows Stormshield's scan of item[0] to interleave before
      item[0] is fully committed in the destination store — hence
      `MAPI_E_NOT_FOUND`.

      **Investigation steps (before a code fix):**
      1. Add `Debug`-level Serilog markers in `ClassifierService.Classify`
         before and after each `_store.MoveItemToFolder(item, …)` call,
         logging item index and elapsed time. Correlate timestamps against
         Stormshield's error timestamp to confirm the interleaving.
      2. Reproduce with exactly 2 emails to find the minimum count.
      3. Verify the Stormshield notification is its own dialog/toast and
         not an exception propagating into our `errors` counter.

      **Candidate fix (implement after investigation confirms hypothesis):**
      After each item's move, release the STA for one COM event cycle
      before processing the next item. In the classify loop in
      `ClassifierService` (or in `RblClassAddIn` if classify is driven
      from the shell), insert an `await Dispatcher.Yield()` (or
      equivalent STA-thread yield). This gives Stormshield's async scan
      time to complete against item[0] while item[0] is still stable in
      the destination store, before we call `GetItemFromID()` for item[1].
      If classify currently runs synchronously, making it `async Task`
      is the prerequisite. No settings change unless a fallback
      `InterItemDelayMs` hidden setting is needed.

### Carried over from v2.2.0.0

- [ ] **Strip the external-sender banner on reply and on classify** —
      implemented in v2.1.0.0 scope; not verified live (no
      external-banner mail on the dev machine). Re-verify on a
      workstation that receives the company banner before closing.
