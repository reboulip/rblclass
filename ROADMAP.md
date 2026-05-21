## Phase 0 — De-risking (1 week)

### Already validated ✅
- Outlook target confirmed: Outlook **32-bit**, Semi-Annual Enterprise
  channel. No additional internal admin restriction on COM add-ins
  (third-party add-in Stormshield is loaded and running). No GPO blocks
  add-ins or macros loading. User has write access to
  HKCU\Software\Microsoft\Office\Outlook\Addins\.
- Internal PKI is available for code signing.

### Remaining tasks
- [ ] Confirm the internal PKI can issue a certificate with the
      **Code Signing EKU** (OID 1.3.6.1.5.5.7.3.3). A generic client
      certificate is NOT sufficient.
- [ ] Verify ClickOnce per-user install works on a real user workstation
      (no admin).
- [ ] Build the "Hello PST" **COM Add-in** POC:
    - Ribbon button (via `IRibbonExtensibility`)
    - Enumerate open PST stores, count items in first folder
    - Open a SQLite database in %LocalAppData%, write/read one row
      (validates that `Microsoft.Data.Sqlite` + the native
      `e_sqlite3.dll` load correctly inside the Outlook process at
      both 32-bit and 64-bit, via the `SQLitePCLRaw.bundle_e_sqlite3`
      both-arch runtime payload)
    - Package as a signed per-user installer (PowerShell script for
      the POC; WiX MSI from Phase 1) and verify it installs without
      admin rights
- [ ] Deploy the POC for validation:
    - Dev-machine smoke test on 64-bit Outlook, Current channel
    - Target-workstation validation on 32-bit Outlook, Semi-Annual
      Enterprise channel — observe for 2-3 days for any EDR alert
      (Stormshield in particular) and any `LoadBehavior` flip from
      3 to 2 in HKCU
- [ ] Go/No-Go gate.

## Phase 1 — Foundations (2-3 weeks)

Goal: clean architecture, no business value yet, but everything in place
to build fast afterwards.

- [ ] Solution skeleton:
    - RBLclass.Core (.NET Standard 2.0) — business logic, no Outlook dep
    - RBLclass.Outlook.Adapter (.NET Framework 4.8) — IMailStore impl over COM
    - RBLclass.AddIn (.NET Framework 4.8) — thin COM-add-in shell
      (IDTExtensibility2 + IRibbonExtensibility)
    - RBLclass.Core.Tests (xUnit)
- [ ] Define core interfaces: IMailStore, IFolderTree, IMailIndex, IClassifier
- [ ] SQLite schema v1 (mails index with FTS5, folders, conversation→folder
      mapping, settings, migrations table)
- [ ] Logging infrastructure (Serilog, rolling file in %LocalAppData%)
- [ ] CI: build + tests on every push (GitHub Actions or Azure DevOps)
- [ ] COM lifetime management utilities (ComWrapper<T> with IDisposable)
- [ ] WiX-based MSI installer skeleton, per-user, Authenticode-signed,
      registers the add-in under
      `HKCU\Software\Microsoft\Office\Outlook\Addins\`

## Phase 2 — Indexing engine (2 weeks)

- [ ] Full initial indexing of all open PST stores, background thread,
      progress reporting, cancellable
- [ ] Persistent index in SQLite, keyed by Outlook EntryID + StoreID
- [ ] Incremental updates via Outlook events (ItemAdd, ItemChange, ItemRemove)
- [ ] Folder tree caching and change tracking
- [ ] Performance target: search returns results in <100ms on 100k mails

## Phase 3 — Core UX (3 weeks)

- [ ] Custom Ribbon (Ribbon XML) with main actions
- [ ] WPF Task Pane with MVVM (CommunityToolkit.Mvvm)
- [ ] Fast folder picker with fuzzy keyword search
- [ ] "Classify selected mail(s)" command — optimistic UI, async Move
- [ ] Folder content preview pane
- [ ] Settings dialog (paths to index, exclusions, hotkeys)

## Phase 4 — Productivity features (2 weeks)

- [ ] Attachment management (list, remove, save)
- [ ] Configurable pop-up on ItemSend (rules: external recipient, keywords...)
- [ ] Keyboard shortcuts for top folders (most used / pinned)

## Phase 5 — Smart features (2 weeks)

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
- [ ] Self-service install page (link to ClickOnce manifest)
- [ ] Support process
- [ ] Monitoring of error logs across the fleet

## Phase 8 — Continuous maintenance (ongoing)

- [ ] Regression test plan executed on every Semi-Annual Enterprise
      channel update (typically January and July). Test matrix: install,
      indexing, classify, search, send-time popup, attachment removal.
- [ ] Watch Microsoft 365 roadmap and deprecation notices for VSTO and
      classic Outlook.
- [ ] Annual review of the migration plan toward "new Outlook" if/when
      that becomes a corporate direction.
