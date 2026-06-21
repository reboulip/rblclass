# RBLclass roadmap — closed history

This is the **archived record** of completed roadmap work. The active plan
(current sprint + open forward-looking phases) lives in
[../ROADMAP.md](../ROADMAP.md). Nothing here is outstanding; it is kept for
provenance and the regression trail.

Reference docs produced from the legacy code analysis:
- [legacy-overview.md](legacy-overview.md) — feature memo of the legacy VBA
  RBLclass.
- [legacy-reimplementation-roadmap.md](legacy-reimplementation-roadmap.md) —
  the sequential, feature-by-feature rebuild plan (Steps 1–10) and the
  approved deviations.
- [parity-and-regression.md](parity-and-regression.md) — the parity checklist
  and regression matrix.

The legacy-reimplementation roadmap (Steps 0–10) is **complete and signed
off**. Phases 1–4 below map onto those steps and are done, modulo the recorded
deviations (Quick Open dropped, attachment management folded into Step 4,
incremental re-indexing landed narrower, reminder-to-front dropped). The
throwaway Phase 0 POC was removed once the product stabilised (2026-06-09).
Product was relabelled **v2.0.0.0** — the first ship of the rewritten add-in,
succeeding the legacy VBA macro.

## Phase 0 — De-risking (done)

### Validated
- Outlook target confirmed: Outlook **32-bit**, Semi-Annual Enterprise
  channel. No internal admin restriction on COM add-ins (Stormshield loads
  alongside). No GPO blocks. User has write access to
  `HKCU\Software\Microsoft\Office\Outlook\Addins\`.
- **"Hello PST" COM Add-in POC built and validated end-to-end on the real
  32-bit target** (ribbon button, PST enumeration, `Microsoft.Data.Sqlite` +
  native `e_sqlite3.dll` at both bitnesses via
  `SQLitePCLRaw.bundle_e_sqlite3`).
- **Per-user install validated.** Unsigned PowerShell installer writing HKCU
  COM registration installs and loads without admin rights, no Stormshield/EDR
  objection, no `LoadBehavior` 3→2 flip.
- **No code signature required to load.** Authenticode signing is optional
  defence-in-depth, not a functional gate.

### Notes on signing & packaging
- ClickOnce is **not used** (VSTO-specific, excluded). Packaging path: unsigned
  PowerShell installer for the POC → WiX per-user MSI for GA, signing optional
  on both.
- Confirming the internal PKI can issue a **Code Signing EKU** cert (OID
  1.3.6.1.5.5.7.3.3) is a nice-to-have, not a blocker.

### Status
- [x] POC built, deployed, validated on the 32-bit target.
- [x] Go/No-Go gate passed.

## Phase 1 — Foundations (done, bar CI)

- [x] Solution skeleton (Step 0, commit `481f0b3`): RBLclass.Core (.NET
      Standard 2.0), RBLclass.Outlook.Adapter (.NET FW 4.8), RBLclass.AddIn
      (.NET FW 4.8), RBLclass.Core.Tests (xUnit).
- [x] Core interfaces: `IMailStore`, `IFolderTree`, `IFolderSearch`,
      `IClassifier`, `ISettingsStore`, `IFolderRepository` (Step 1 —
      [[step1-architecture-decisions]]). `IMailIndex` deferred to Phase 5.
- [x] SQLite schema v1: `Folders`, `Settings`, `SchemaVersion`. Mail FTS5
      index correctly **not** built — legacy parity never needed it.
- [x] Logging (Serilog rolling file sink, `%LocalAppData%\RBLclass\logs\`).
- [ ] **CI — still open.** No `.github/workflows` yet; the `RBLclass.Core`
      tests run and pass locally on every change. Lowest-risk remaining
      Phase 1 item.
- [x] POC installer promoted to `Install-RblClassAddIn.ps1` /
      `Uninstall-RblClassAddIn.ps1` (`.claude/skills/make-release/`).
- [x] COM lifetime utilities (`ComRef<T> : IDisposable`).
- [x] WiX MSI installer — built out fully, per-user, validated by
      install/uninstall on the dev machine (commit `0b17715`, rebuilt as
      `RBLclass-2.0.0.0.msi`). The PowerShell kit and MSI coexist until the
      MSI is also validated on the 32-bit target.

## Phase 2 — Folder indexing engine (done, two deviations)

- [x] Initial folder-tree walk of all open PST stores on **first run only**,
      background thread, progress + cancellation (`FolderIndexService`,
      Step 1).
- [x] **Persistent** folder index in SQLite, keyed by (StoreId, EntryId).
- [x] *(narrower than planned)* Targeted re-index on user-triggered sub-folder
      creation (`IFolderTree.ReindexStore`, landed in Step 4). Passive
      `FolderAdd`/`FolderRemove`/`FolderChange` subscription did **not** land;
      no drift surfaced in dev validation.
- [x] Configurable store/folder exclusions (`FolderExclusionPolicy`, Step 1).
- [x] *(configurable option, not a fixed mode)* Folder keyword search
      (`FolderSearchService`/`IFolderSearch`): AND across keywords, case/accent
      insensitive, collapse/expand, result cap. Match mode is a user setting —
      defaults to word-prefix with substring opt-in (Step 9,
      [[folder-search-match-mode]]); later flipped to substring-default in
      v2.1.0.0.
- [x] Folder search < 100 ms verified live on the dev tree.

## Phase 3 — Core UX (done, one drop + one stack deviation)

- [x] Custom Ribbon (Ribbon XML).
- [x] *(approved deviation)* WPF Task Pane with hand-rolled minimal MVVM
      (`RBLclass.AddIn/Mvvm`), bridged via `ElementHost` into a ComVisible
      WinForms CTP host — **not** CommunityToolkit.Mvvm (its 8.x dep graph
      breaks SQLite under a binding-redirect-less COM host, commit `f66e2b1`,
      [[com-addin-no-binding-redirects]]).
- [x] Fast folder picker over IFolderSearch (Step 3).
- [x] "Classify selected mail(s)" — multi-destination, copy-per-destination,
      attachment handling on copy ([[classify-attachments-on-copy]]) (Step 4).
- [x] Inline sub-folder creation + targeted re-index (Step 4).
- [x] Conversation widening + task-completion guard (Step 6).
- [x] ~~Quick Open "jump to last filed folder" (Step 5)~~ — **DROPPED, not
      deferred** ([[step5-quick-open-dropped]]).
- [x] Settings pane backed by the SQLite `Settings` table (Step 9: typed
      `Settings` in Core + WPF `SettingsWindow`; eleven options —
      [[settings-always-persistent]]).

## Phase 4 — Productivity features (done, one drop)

- [x] *(landed inside Step 4)* Attachment management — standalone "Remove
      attachments" + the classify-time toggle, **filed-copy-only** semantics
      ([[classify-attachments-on-copy]]).
- [x] Configurable `ItemSend` guards (Step 7, commit `e87a386`):
      `ForgottenAttachmentGuard` + `ExternalRecipientGuard` (AddressEntry type
      + internal-domain allowlist), both backed by Settings.
- [x] Sent-item triage prompt — Class / Delete / Move-to-Inbox / Leave,
      whole-conversation, re-entrancy guard (`SentItemTriageWindow`/`ViewModel`,
      Step 8, commit `08226dc`). *(Reworked into a dropdown in v2.1.0.0.)*
- [ ] ~~Keyboard shortcuts for top folders~~ — **not built**, not part of
      legacy parity. Revisit only on pilot demand.

## Phase 5 — Mail content index + smart features (deferred indefinitely)

> Steps 1–10 reached full legacy parity without any of this — the legacy tool
> never indexed mail bodies. Re-scope against real pilot feedback rather than
> building ahead of demand.

- [ ] Deferred mail index (`IMailIndex`): SQLite FTS5 over subject/body, keyed
      by (StoreId, EntryId), incremental via mail item events. Target <100 ms
      on 100k mails.
- [ ] Conversation-based auto-classification suggestion (local rules).
- [ ] `IClassificationSuggester` abstraction.
- [ ] Optional Azure OpenAI implementation behind a feature flag.
- [ ] Response caching in SQLite.

## v2.1.0.0 — Pilot-feedback UX & fixes (shipped)

First feedback wave from operational use. Theming/layout foundations first,
then isolated fixes, then design-heavy redesigns.

### A. Theming & layout foundations
- [x] **Theme-aware task-pane & dialog styling.** Read the Office UI theme
      (White / Colorful / Dark Gray / Black), apply a matching palette across
      the pane and every modal; fall back to the Windows app theme when Office
      is "Use system setting". Centralised in a shared WPF `ResourceDictionary`.
- [x] **Long folder-path display in the narrow pane.** Left-truncate with a
      leading ellipsis (show the leaf), full path on hover.

### B. Isolated correctness & UX fixes
- [x] **"Contains" as the default folder-search match mode**
      ([[folder-search-match-mode]]); word-prefix becomes opt-in.
- [x] **Classify into a PST store's root node** — store root is a selectable
      destination; adapter supports `Move` into the root.
- [x] **Freeze the pane while a classify/move runs (anti-double-fire)** — a
      re-entrancy guard in `ClassifyViewModel.DoClassify`.
- [x] **Process encrypted conversation siblings correctly (Stormshield)** —
      skip but surface a notice; `GetConversationSiblings` returns a
      `ConversationSiblings` (processable + skipped-encrypted subjects).
      *(Unit-tested only; not verified live — no S/MIME mail on the dev
      machine.)*
- [x] **Position modal windows on the active Outlook monitor** — dialogs owned
      by the Outlook main window (`DialogPlacement.OwnByOutlook`),
      `CenterOwner`. *(Not verified live — single-monitor dev machine.)*
- [x] **Pick up folders created/renamed in Outlook, via a manual Refresh** — a
      "Refresh folders" ribbon button re-walks live stores
      (`IFolderTree.WalkAndPersist`). Live folder events rejected (COM-lifetime
      sprawl); periodic timer walk rejected (wasteful).

### C. Design-heavy redesigns
- [x] **Unify the two pane modes into one always-open task pane** built around
      a single folder-search list (single-click toggles selection,
      double-click files immediately, per-row button reveals in the explorer).
      Shipped as new `MainPaneView` + `MainPaneViewModel` + `NamePromptWindow`;
      old `FolderSearchView`/`ClassifyView` + `PaneMode` deleted
      ([[prefer-isolated-new-ui-over-retrofit]]).
- [x] **Rework sent-item triage into a dropdown setting** — Move-to-Inbox /
      Delete / Leave / "Let me choose"; fixed value applies with no modal,
      "Let me choose" shows the modal. Removed the "Class" action and the
      whole-conversation checkbox; settings migration from the old boolean.
- [x] **Strip the external-sender banner on reply and on classify** — learns
      the banner from a user-taught sample, derives a signature, strips
      matching blocks; on reply/forward (toggle) and at classify time (the
      filed copy). *(Core matcher unit-tested; Outlook wiring NOT verifiable
      here — no external-banner mail.)*

## v2.2.0.0 — Second feedback wave (shipped)

Eight items from continued operational use. Search-correctness fixes first,
then the classify-engine rework (the Stormshield move-based fix) that Undo and
Auto-class build on, then feature work.

### A. Search correctness & responsiveness
- [x] **Fix special characters in folder search ("R&D").** A token with a
      non-alphanumeric char falls back to substring containment.
- [x] **Minimum query length + typing debounce.** `MinSearchLength` (default
      2) and `SearchDebounceMs` (default 200, clamped 0–2000).
- [x] **Ctrl toggles "List every matching folder" from the search box** (Ctrl
      alone; Ctrl+key combos unaffected).
- [x] **"Open folders in a new window" becomes settings-only.**
- [x] **Select all text when the search box gains focus by click.**

### B. Classify engine rework
- [x] **Stop provoking Stormshield's `MAPI_E_NOT_FOUND` on classify** —
      move-based classify. KeepCopy off: copies only for destinations 1…n−1,
      original **moved** to the last (single destination ⇒ pure `Move`, no
      transient copy, no Deleted Items churn); `IMailStore.MoveItemToFolder`
      added. KeepCopy on: unchanged. Sent-item "Move to Inbox" rides the same
      path. **Encrypted mail (`IPM.Note.SMIME`) is never stripped of
      attachments** ([[never-strip-encrypted-attachments]]). Opt-in
      `ClassifySafetyCopy` (default off) restores the old Deleted-Items copy,
      taken from the moved item at its destination (no transient — no
      Stormshield race). *(Move/copy/strip/triage verified live; the
      Stormshield error itself and the encrypted-skip NOT verifiable here.)*

### C. Filing features
- [x] **Accumulate destinations across successive searches** — a selection set
      (StoreId+EntryId) survives re-searches; a chip strip above Classify
      (per-chip ✕ + clear-all); clears after a successful classify.
- [x] **Undo the last filing action** — `ClassifierService.Classify` returns
      an undo plan (moves with source folder, copies created, flags set,
      history rows); a single-slot Undo button reverses it. Stripped
      attachments can't be restored; Undo says so.
- [x] **Auto-class** — schema **v2** adds `ClassificationHistory`
      (ConversationKey, DestStoreId, DestEntryId, WhenUtc), appended on every
      classify. Files each mail to its conversation's most recent recorded
      destination(s), validated against the current folder index. A small
      "Auto-class" button in the pane; filed destinations shown in the results
      list; outcomes in the status line.

## v2.4.0.0 — shipped items (third feedback wave)

The quick wins and pane/visual redesigns that shipped before the bug-fix and
new-feature work documented in the active [../ROADMAP.md](../ROADMAP.md).

### A. Quick wins
- [x] **A1. Clear the search field after a successful classify** —
      `ClearQuerySilently()` in `MainPaneViewModel` writes the backing field
      and raises `PropertyChanged` without re-searching; displayed results stay
      until the next keystroke. Also dropped the v2.2 select-all-on-focus
      behaviour.
- [x] **A2. Async folder index refresh with pane colour indicator** — removed
      the end-of-refresh MessageBox; `WalkAndPersist` runs on the STA with
      await-yielding, SQLite writes on the thread pool; a coloured dot in the
      pane header (red = absent/empty, yellow = indexing, green = ready).
      `IFolderIndexService.IndexStatus` (`NotFound`/`Indexing`/`Ready`) bound
      via `MainPaneViewModel`. Tooltips localized.
- [x] **A3. Pane dock position persistence** — subscribe to
      `_ctp.DockPositionStateChange`, persist `PaneDockPosition` (default
      `msoCTPDockPositionRight`); `EnsureTaskPane()` reads it before showing.
      Pane width not persisted (no `WidthChanged` event).

### B. Pane interaction redesign
- [x] **B1. Collapse all checkboxes behind an "Options" toggle** — the five
      per-action checkboxes (AllResults, KeepCopy, RemoveAttachments,
      WidenConversation, StripBanner) hide behind a collapsible panel; expand
      via the Options button or Tab from the search box; collapse on toggle or
      a successful classify, resetting checkboxes to their persisted defaults.
      `MainPaneViewModel.IsOptionsExpanded`; `Pane_Options` localized.

### C. Search results display
- [x] **C1. Per-row expandable folder path** — new `FolderPathDisplay`
      (Inline / Hierarchical, default Inline). Hierarchical splits `FullPath`
      on `\` into `PathSegments`, rendered as an indented `StackPanel`
      (12 px/depth) with the `↗`/`+` buttons only on the leaf row; a
      `DataTemplateSelector` switches templates. Settings toggle in EN/FR/DE;
      xUnit over `PathSegments` splitting.
