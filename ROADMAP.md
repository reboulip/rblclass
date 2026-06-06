# RBLclass roadmap

Reference docs produced from the legacy code analysis — read these for
the per-feature detail this phase plan schedules:
- [docs/legacy-overview.md](docs/legacy-overview.md) — feature memo of
  the legacy VBA RBLclass.
- [docs/legacy-reimplementation-roadmap.md](docs/legacy-reimplementation-roadmap.md)
  — the sequential, feature-by-feature rebuild plan (Steps 1–10) and the
  approved deviations.

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

- [ ] Solution skeleton:
    - RBLclass.Core (.NET Standard 2.0) — business logic, no Outlook dep
    - RBLclass.Outlook.Adapter (.NET Framework 4.8) — IMailStore impl over COM
    - RBLclass.AddIn (.NET Framework 4.8) — thin COM-add-in shell
      (IDTExtensibility2 + IRibbonExtensibility)
    - RBLclass.Core.Tests (xUnit)
- [ ] Define core interfaces: IMailStore, IFolderTree, IFolderSearch,
      IClassifier, ISettingsStore. (IMailIndex is **deferred** to Phase 5
      — legacy parity indexes folders only; see Phase 2.)
- [ ] SQLite schema v1: `Folders` (StoreId, EntryId, ParentEntryId,
      Name, FullPath, IsLeaf), `Settings`, `SchemaVersion` migrations
      table. The **mail FTS5 index is NOT in v1** — it lands in Phase 5
      with the smart features that need it.
- [ ] Logging infrastructure (Serilog, rolling file in %LocalAppData%)
- [ ] CI: build + Core unit tests on every push (GitHub Actions or
      Azure DevOps)
- [ ] Promote the POC PowerShell installer to a reusable
      `Install-RBLclass.ps1` / `Uninstall-RBLclass.ps1` so every
      subsequent increment is loadable on the dev workstation for manual
      validation (see Testing & validation strategy above).
- [ ] COM lifetime management utilities (ComRef<T> with IDisposable)
- [ ] WiX-based MSI installer skeleton, per-user, registers the add-in
      under `HKCU\Software\Microsoft\Office\Outlook\Addins\` (signing
      optional — not required for load; keep the POC PowerShell
      installer working in parallel until the MSI is validated on the
      target)

## Phase 2 — Folder indexing engine (2 weeks)

Legacy parity needs a **folder** index only — the legacy tool never
indexed mail bodies (see [docs/legacy-overview.md](docs/legacy-overview.md)).
This phase delivers Steps 1–2 of the reimplementation roadmap. The mail
content index is deferred to Phase 5.

- [ ] Initial folder-tree walk of all open PST stores on **first run
      only**, on a background thread, with progress + cancellation
      (replaces the legacy modal "please wait" form).
- [ ] **Persistent** folder index in SQLite, keyed by (StoreId, EntryId).
      On every subsequent start the tree is **loaded from SQLite, not
      re-walked from Outlook** — the index rests in the database between
      sessions.
- [ ] Incremental freshness via Outlook **folder** events
      (FolderAdd / FolderRemove / FolderChange) + a cheap background
      reconcile; targeted re-index of a single subtree on sub-folder
      creation. No full rebuild on startup.
- [ ] Configurable store/folder exclusions (replacing the legacy
      hard-coded FR/EN substrings like "Dossiers publics" / "G-FRA").
- [ ] Folder keyword search (IFolderSearch): AND across keywords,
      **substring**, case- and accent-insensitive, with collapse/expand
      and a result cap — matched over the SQLite-hydrated tree, **not
      FTS5** (FTS5 tokenising would break substring hits like
      "PJ" → "ProjetJuridique").
- [ ] Performance target: folder search < 100 ms on a large tree
      (thousands of folders).

## Phase 3 — Core UX (3 weeks)

Delivers reimplementation roadmap Steps 3–6 (open-folder, classify,
quick-open, conversation widening + task guard) and Step 9 (settings).

- [ ] Custom Ribbon (Ribbon XML) with main actions
- [ ] WPF Task Pane with MVVM (CommunityToolkit.Mvvm)
- [ ] Fast folder picker over IFolderSearch (open-a-folder, Step 3)
- [ ] "Classify selected mail(s)" command — multi-destination,
      copy-per-destination, optimistic UI, async Move (Step 4)
- [ ] Inline sub-folder creation + targeted re-index (Step 4)
- [ ] Conversation widening + task-completion guard (Step 6)
- [ ] Quick Open "jump to last filed folder" affordance (Step 5)
- [ ] Settings pane backed by the SQLite `Settings` table — replaces the
      legacy positional `options.txt` (Step 9)

## Phase 4 — Productivity features (2 weeks)

Delivers reimplementation roadmap Steps 5, 7, 8.

- [ ] Attachment management — standalone "remove attachments" + the
      classify-time toggle (Step 4/5)
- [ ] Configurable ItemSend guards: forgotten-attachment (configurable
      keyword list) and external-recipient (AddressEntry type +
      configurable internal-domain allowlist, **not** the legacy
      `CORPMAIL` substring) (Step 7)
- [ ] Sent-item triage prompt — class / delete / move-to-inbox / leave,
      whole-conversation, with a clean re-entrancy guard (Step 8)
- [ ] Keyboard shortcuts for top folders (most used / pinned)

## Phase 5 — Mail content index + smart features (2 weeks)

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

- [ ] Deploy to 5-10 pilot users
- [ ] Telemetry (opt-in): performance metrics, error reports
- [ ] Feedback loop, bug fixes
- [ ] Documentation: user guide, troubleshooting

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
