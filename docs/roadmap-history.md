# RBLclass roadmap — closed history

This is the **archived record** of completed roadmap work. The active plan
(current sprint + open forward-looking phases) lives in
[../ROADMAP.md](../ROADMAP.md). Nothing here is outstanding, except where a
section note points back to ROADMAP.md (items still awaiting a real-target
check); it is kept for provenance and the regression trail.

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

## Phase 6 — Pilot rollout (done)

- [x] Documentation: trilingual (EN/DE/FR) self-contained single-file HTML
      install guide (`docs/installation-guide.html`, commit `beee310`). **Still
      needs** its `[contact / channel]` placeholder filled in (all three
      languages) before it is emailed to pilots.
- [x] Packaging: product relabelled **v2.0.0.0**, rebuilt as
      `RBLclass-2.0.0.0.msi` (`installer\Build-Installer.ps1`), validated by
      install/uninstall on the dev machine.
- [x] **Open `develop` → `main` PR (#2)** — gating merge on re-running the
      [docs/parity-and-regression.md](docs/parity-and-regression.md) matrix on
      the real 32-bit target, then tagging `v2.0.0.0` on `main` and packaging
      with `/make-release`.
- [x] Deploy to 5–10 pilot users.
- [x] Telemetry (opt-in): performance metrics, error reports.
- [x] Feedback loop, bug fixes.

## Phase 7 — General availability (done)

- [x] Communication plan.
- [x] Self-service install page (link to the per-user MSI on the internal HTTPS
      share).
- [x] Support process.
- [x] Monitoring of error logs across the fleet.

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

## v2.4.0.0 — Third feedback wave (shipped)

Quick wins and pane/visual redesigns (A–C), then the classify
performance/freeze fix (D), classify-after-send (E), and the
attachment-disposition capability (F).

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

### D. Classify reliability & performance
- [x] **D1. Stormshield `MAPI_E_NOT_FOUND` on multi-item classify — cosmetic.**
      Folded into D2: the inter-item pump-yield gives Stormshield's async scan a
      stable window after each `Move`. Cosmetic only (classify succeeds, mail
      lands); manifests with multiple items selected. Re-verify on the
      Stormshield 32-bit target alongside D2; optional hidden `InterItemDelayMs`
      fallback if a residual remains.
- [x] **D2. Responsive classify (no UI freeze).** Classify ran synchronously on
      the STA UI thread, blocking the message pump for the whole run (Outlook
      appeared hung). Reworked to an async path that yields the STA pump between
      items (`Dispatcher.Yield`/`DispatcherFrame`) so Outlook repaints and stays
      interactive; per-item progress in the status line
      (`Status_Classify_Progress`); non-COM work (`ClassificationHistory`
      writes, undo-plan persistence) dispatched to the thread pool; the COM
      moves stay on the STA. Core stays UI-free via an injected yield callback;
      `IsBusy` anti-double-fire preserved. xUnit over the progress/yield
      callbacks and undo-plan invariance.

### E. Classify after send
- [x] **E1. Re-offer classification of a just-sent mail via Move-to-Inbox.**
      When sent-item triage resolves to "Move to Inbox", the moved
      `MailItemRef` is pinned into `MainPaneViewModel` as the active classify
      target (overrides `GetSelectedItems` for the next classify; clears after
      that classify or on a real selection change) and the pane is revealed.
      New setting `ClassifyAfterMoveToInbox` (default on). The OM has no reliable
      API to set `Explorer.Selection`, hence the pin approach rather than
      selecting the item in the message list.

### F. Attachment disposition & filesystem favourites
- [x] **F1. Keyword-searchable "favourite folders" filesystem index.** Settings
      editor to add/remove favourite Windows directories
      (`AttachmentFavoriteFolders`); recursive **path-only** indexing with a
      depth/count cap (32-bit memory budget); persisted in SQLite **schema v3**
      (`FavoriteFolders`) and searched with the same
      `FolderSearchService`/`FolderMatchMode` matcher. Directory walk behind an
      injected `IDirectoryScanner`/`IFavoriteFolderSource`. xUnit over keyword
      matching, path-segment splitting, caps, empty/duplicate paths.
- [x] **F2. Per-attachment disposition modal.** New setting
      `AttachmentRemovalMode` = **Modal** (default) / **DeleteSilently**. When
      Modal + remove-attachments, classify shows a modal listing each attachment
      (filename, size, source subject) with a per-row **Delete** or **Save to…**
      (folder picker over the F1 favourites index + browse fallback). Encrypted
      mail (`IPM.Note.SMIME`) excluded entirely and reported
      ([[never-strip-encrypted-attachments]]). On confirm: save each (collision
      resolved with ` (1)`, ` (2)`…), strip the chosen attachments, apply the F3
      label — on the filed copy / moved item ([[classify-attachments-on-copy]]).
      Adapter gains `GetAttachments` / `SaveAttachmentToFile`; `AttachmentInfo`
      is an Outlook-free Core type. **Fails safe** when a save is blocked
      (Windows Controlled Folder Access): nothing stripped, run reports a
      failure. Confirm CFA allow-listing of `outlook.exe` on the 32-bit target.
- [x] **F3. "Former attachments" label on the filed mail.**
      `AttachmentLabelLocation` = **InfoBar** / **Body**. Body path shipped: an
      HTML note block appended to `HTMLBody` listing each former attachment and
      its disposition ("Saved to `<path>`" / "Deleted on `<date>`"), reusing the
      banner-rewrite pattern. InfoBar (`OlkInfoBar`) is a form-region control
      with no durable per-item label for stored items — **recorded as deferred**,
      Body is the fallback. Pure-Core `AttachmentLabelFormatter` with
      `_One`/`_Other` plurals. xUnit over saved/deleted/mixed compositions and
      single-vs-multiple (plural) rendering.

## v2.4.1.0 — Attachment & indexing bug fixes (shipped)

- [x] **A1.** Re-index button fully offloads PST enumeration to a background
      thread so the Outlook UI stays responsive during manual reindex, colour
      indicator cycling to yellow. [#6]
- [x] **A2.** Attachment-disposition modal shown during auto-classify when
      "remove attachments" is on, matching the manual-classify flow. [#7]
- [x] **A3.** Discrete status-area notice instead of the modal when classifying
      encrypted mail (which always retains its attachments). [#8]

## v2.5.0.0 — Attachment refinements & UX polish (shipped)

### A. Status & notifications
- [x] **A1.** Single status-area notification replaced by a rolling in-session
      log capped at ~5 entries, so short-lived messages stay visible. [#9]

### B. Attachment enhancements
- [x] **B1.** "Keep" disposition per attachment row in the removal modal,
      leaving chosen attachments on the filed copy untouched. [#5]
- [x] **B2.** Inline/embedded images auto-excluded from the attachment-removal
      flow, so only true detached files are presented for disposition. [#4]

### C. Packaging
- [x] **C1.** MSI detects and removes a previously installed version before
      laying down the new one, eliminating the manual uninstall step. [#10]

## v2.6.0.0 — UI refresh & classify depth (shipped)

> Archived 2026-11-21. B3 and C1 were never re-checked on the real target; they were dropped from tracking that day (issues #19 and #12 closed).

### A — UI & ribbon
- [x] **A1.** Ribbon UI overhaul: add a Settings icon, remove the "Remove Attachments" and "Index Status" buttons, add an About button (version / author / tech info), and an easter egg — clicking the hidden item in About replaces the Auto-Class button label with a pig icon. [#15]
- [x] **A2.** Increase the prominence of the Classify button in the task pane. [#17]
- [x] **A3.** Add an auto-expand toggle and a visual row separator to search results. [#13]

### B — Classify & attachments
- [x] **B1.** Extend the auto-classify conversation-history retention window and expose it as a configurable setting. [#16]
- [x] **B2.** Add a "use the same save directory" checkbox (checked by default) to the multi-mail attachment disposition UI. [#18]
- [ ] **B3.** Add diagnostics for external-sender banner detection, improve reliability, and add fine-grained configuration; closes the open carried-over verification item. [#19] - shipped, pending tests

### C — Meeting items
- [ ] **C1.** Support classifying MeetingItem objects from the selection, enabled via an opt-in toggle in Settings. [#12] - shipped,pending tests

## v2.7.0.0 — feedback wave: shortcuts, sent-items, docs (shipped)

*(v2.5.1.0 — diagnostic PERF logging — and v2.5.2.0 — classify-time freeze
fix: deferred post-classify refresh + mid-batch selection suppression —
shipped between v2.5.0.0 and this wave without their own roadmap sections.)*

### A — Keyboard shortcuts
- [x] **A1.** Add Ctrl+Z as a keyboard shortcut for the Undo action. [#30]
- [x] **A2.** Add keyboard shortcuts to open the RBLclass pane, focus the search field when the pane is open, and open the searched folder (distinct from Enter, which classifies); specific bindings chosen with the user; add a startup option to auto-show the pane when Outlook opens. [#22] [#31]

### B — Sent-items handling
- [x] **B1.** Clarify the post-send classify dialog labels and lay the buttons out vertically so long labels render clearly in every language. [#24]
- [x] **B2.** Add a configurable default archive folder into which sent messages are moved automatically. [#23]

### C — Search result display
- [x] **C1.** Add a Settings option for the expanded search-result path display, letting the user choose between the current one-line-per-folder layout and a multi-line full-path layout; default stays the one-liner. [#26]

### D — Classify
- [~] **D1.** ~~Add an opt-in (off by default) to keep only the latest message of a conversation in the target folder on classify — deleting older ones but never those with attachments; use Outlook's built-in conversation clean-up if suitable.~~ **DEFERRED/DROPPED (v2.7.0.0):** Outlook exposes no programmatic conversation clean-up API (`MailItem.GetConversation()` has no `CleanUp`), and the only way to trigger the real feature — `CommandBars.ExecuteMso("CleanUpConversation")` — acts on the active explorer's selection/folder and respects the user's global clean-up settings, so it cannot be reliably targeted at a just-classified conversation without fragile UI automation. A prototype that manually deleted older siblings by date was rejected (it lacks Outlook's redundancy semantics). Recommend users apply Outlook's native **Clean Up Conversation** instead. [#27]
- [x] **D2.** Add an opt-in setting (off by default) to mark classified mails as read automatically on classify. [#32]

### E — Update notification (investigation)
- [x] **E1.** Investigate a startup check that detects a newer RBLclass release and surfaces an orange in-pane update banner; account for corporate-proxy constraints and the no-network-by-default posture — defer or drop if it proves too complex. [#25] **Implemented** as an opt-in (off by default), no default URL: a single startup GET to a user-configured HTTPS version marker via `HttpClient` (honours the system proxy), showing an orange in-pane banner when a newer version is available; dismissal snoozes it for 7 days. This introduces the first outbound network call; the Phase 9 "no outbound network calls" posture is documented accordingly in `docs/security.md` (opt-in, off by default, no data sent beyond the GET).

### F — Documentation
- [x] **F1.** Produce a simplified end-user usage guide (features, per-setting behaviour, how to suggest improvements) in both multi-language HTML and MediaWiki format (one .txt file per supported language). [#21]

## v2.7.1.0 — code-review remediation sprint (shipped)

Fixes from the 2026-07-03 full code review (diff review of the v2.7.0.0 sprint
+ whole-codebase review; 20 ranked findings + verified mentions). Grouped so
each item is one coherent, testable change.

### A — Data integrity (classify & attachments)
- [x] **A1.** Mark-as-read hardening in `OutlookMailStore`: extract the duplicated
      `UnRead = false; Save()` block from `CopyItemToFolder`/`MoveItemToFolder`
      into one shared helper that (a) skips `IPM.Note.SMIME*` items via the same
      `StartsWith(OrdinalIgnoreCase)` guard every other content-mutating path
      uses (Save on signed/encrypted mail risks corruption — the
      `ClassifyRequest.MarkAsRead` doc comment already promises "S/MIME-safe"),
      and (b) also marks `MeetingItem` objects read (they expose `UnRead`) so the
      setting applies to meeting classify too.
- [x] **A2.** Attachment disposition dialog: "use the same save directory" must
      no longer override rows the user explicitly set to Delete/Keep —
      `ApplySameDirectory` applies the chosen directory only to rows still in
      their default state, never flipping an explicit non-SaveTo choice to SaveTo.
      **Extended during verification:** the default-checked radio per row is now
      *Save to* (non-destructive default) instead of *Delete*, per user feedback.
- [x] **A3.** `OutlookMailStore` small integrity fixes: (a) when keep-copy filing
      fails (`Copy()` succeeded, `Move()` threw), best-effort delete the
      transient copy before rethrowing so no orphaned duplicate accumulates in
      the source folder; (b) in `AddSiblingIfInScope`, replace the exact
      `== "IPM.Note.SMIME"` check with the `StartsWith("IPM.Note.SMIME",
      OrdinalIgnoreCase)` guard used everywhere else, so signed/encrypted
      subtypes (e.g. `.MultipartSigned`) are reported as skipped siblings.

### B — Exception safety & result accounting
- [x] **B1.** Guard every unguarded UI entry point so no exception can escape to
      the dispatcher inside Outlook: the Ctrl+Enter `OpenHighlighted()` branch in
      `QueryBox_KeyDown`, `AutoClass_Click`, `Undo_Click`, and the update-banner
      `BeginInvoke` action in `ScheduleUpdateCheck` (try/catch + `Log.Error`,
      matching `Classify_Click`). Also make `RblClassTaskPaneHost.EnsureView`
      recoverable: if ViewModel construction throws, reset `_view` to null so a
      later show retries instead of leaving a permanently blank pane.
- [x] **B2.** Fix classify strip-failure accounting: wrap the strip-all
      `_store.RemoveAttachments(filedRef)` call (and `MarkTaskComplete`) in
      try/catch like the sibling per-attachment calls, so a strip failure is
      logged and reported as a strip failure instead of double-counting the same
      action as both filed and errored; guard the `mail.Attachments`/`.Count`
      enumeration inside `OutlookMailStore.RemoveAttachments` with the file's
      `Safe()` pattern.
- [x] **B3.** Surface partial auto-classify failures: carry the inner
      `ClassifyResult.Errors`/`EncryptedStripSkips` into `AutoClassifyResult`
      and reflect them in the auto-class status line, so "filed to A, failed at
      B" is no longer reported as a clean success.
      **Extended during verification (two root-cause fixes):** (a) auto-class
      now skips destinations the mail already sits in (Outlook rejects a
      same-folder move with 0x80020009) and reports them as "already in its
      remembered folder"; (b) the post-send triage's move-to-inbox /
      move-to-archive no longer record classification history
      (`ClassifyRequest.RecordHistory = false`) — they had been overwriting
      every conversation's remembered folder with the inbox/archive, silently
      breaking auto-class for those conversations.

### C — Sent-item triage hardening
- [x] **C1.** Triage handler fixes (one change, three parts): (a) set the
      re-entrancy suppression before `ShowDialog()` (reset in `finally`) so a
      second sent item arriving during the open triage dialog cannot stack a
      second modal; (b) surface stale archive-folder failures — Settings
      re-resolves the configured folder's liveness so the existing "no longer
      found" placeholder actually shows, and a failed MoveToArchive gives
      user-visible feedback (pane status) instead of a log-only silent leave;
      (c) pass the already-loaded `Settings` into `ApplySentItemTriage` instead
      of reloading a second full snapshot in the MoveToArchive case.

### D — Startup & pane behaviour
- [x] **D1.** Add the persisted setting (promised by v2.7.0.0 A2) gating the
      startup auto-show of the pane — default ON to preserve current behaviour,
      opt-out in Settings → General; localized in all three languages.
- [x] **D2.** First-run folder walk must yield between stores
      (`Dispatcher.Yield`, like the manual refresh path `RunFolderWalkAsync`)
      so Outlook stays responsive during the initial index on multi-PST profiles.
- [x] **D3.** Propagate mid-session settings changes to the pane immediately:
      re-seed the pane option fields (keep copy, remove attachments, widen
      conversation, strip banner, mark as read) and the path-display default
      when the Settings window persists a change, so the very next classify uses
      the new values (the user guide promises "applied immediately"). While
      there, delete `ResolvePathDisplayMode`'s duplicated migration logic in
      favour of the parsed value already in the `Settings` snapshot loaded by
      `Refresh()`, and move the orphaned `AddResultRow` doc comment back to its
      method.

### E — Settings window & persistence
- [x] **E1.** `SettingsWindow`: flush the focused TextBox's pending
      LostFocus-triggered binding before closing, so Enter (IsDefault) or Esc
      (IsCancel) no longer silently discards the last edit.
- [x] **E2.** `SqliteSettingsStore` hardening: set `PRAGMA journal_mode=WAL`
      like the other two stores (via one shared connection-open helper used by
      all three); add batch `GetAll()`/`SetMany()` so `Settings.Load`/`Save`
      use one connection instead of ~30; make the `Settings` table DDL
      single-source so the copy in `SqliteFolderRepository.ApplyV1` cannot
      silently drift (it is dead code today).
- [x] **E3.** Hot-path costs on the search path: stop calling a full
      `Settings.Load` per keystroke just to read `SearchDebounceMs` (cache it,
      invalidated on settings save via D3's propagation), and cache each
      folder's normalized path + word split (computed at index load/reindex)
      so `Search()` no longer re-normalizes the entire tree per query.

### F — API consistency, small fixes & docs
- [x] **F1.** Consolidate the classify per-action toggles (`keepCopy`,
      `removeAttachments`, `markTasksComplete`, `safetyCopy`, `markAsRead`, …)
      into a `ClassifyOptions` object accepted by `ClassifyRequest` and the
      `IMailStore` filing methods, ending the per-sprint positional-bool growth
      (10-parameter ctor today); update all call sites and test fakes.
- [x] **F2.** Add the missing "mark as read" checkbox to the pane's Options zone
      as the fifth per-action override (seeded from settings, reset after each
      classify, like the other four) so the shipped user guide's description of
      five overrides becomes accurate.
- [x] **F3.** Favourite attachment folders: validate on add that the chosen
      directory path contains no `;` (the settings list separator) and reject
      with a localized message, so a legal-but-unusual NTFS name cannot corrupt
      the favourites list on reload.

### G — Release stamp
- [x] **G1.** Bump `AssemblyVersion`/`AssemblyFileVersion` (AssemblyInfo.cs) and
      `release.config.json` to **2.7.1.0** with the sprint summary comment
      (third-field bump so the MSI MajorUpgrade sees it as newer).

## v2.8.0.0 — feedback wave: pane actions, sent-items, classify reliability (shipped)

### A — Pane actions
- [x] **A1.** Add a "save as" button to the pane, next to the Auto-Class button, letting the user save the selected email(s) as files, reusing the attachment-disposition modal's directory picker for the target folder. [#40]
- [x] **A2.** Add a manual attachment-removal button: opens the attachment-disposition modal for the selected email(s) without a classify move, so keep/delete/save-to can be applied standalone. [#38]
- [x] **A3.** In the standalone attachment-management modal (A2's 📎 button), add a "remove saved attachments from the email" checkbox that controls whether a "Save to" attachment is stripped from the email after saving (ON = save then strip, OFF = save but keep it on the email). The checkbox default comes from a new setting in Settings → Attachments. S/MIME items are never stripped. The classify-time modal is unaffected.

### B — Sent-items handling
- [x] **B1.** Add a "delete" button to the post-sent classify notification, to allow one-click deletion of the sent-and-moved-to-Inbox copy. [#39]

### C — Classify reliability
- [x] **C1.** Fix auto-class + attachment-modal conflict: auto-class must check for a valid target folder in the conversation history *before* opening the attachment-disposition modal, so the modal never opens for an email it cannot process; when multiple emails are selected, only show attachments for the ones that will actually be processed; surface a clearer (non-modal) message when auto-class cannot process an element. [#36]
- [x] **C2.** Investigate and fix the classify silent-failure/freeze reported on v2.6.0.0: ~30s "Classement en cours" then no move, reproducible after reinstall/reboot; log attached on the issue. Reporter notes meeting-invite classify is unaffected — only regular mail. [#35] *(Shipped an error-handling fix for the conversation-index timeout; the reporter's ~30 s stall persisted and was re-scoped in v2.9.0.0 B3 once the repro narrowed to conversation widening.)*

### D — Logging audit & rework
- [x] **D1.** Demote the per-item PERF/D1 BEGIN-END/selection-change chatter (`OutlookMailStore`, `RblClassAddIn`, `MainPaneViewModel`) from Information to Debug, gated by a new persistent `DiagnosticLogging` setting (Settings → Diagnostics, default off, takes effect after restart) that raises the minimum log level back to Debug on demand. The classify/auto-class outcome summary line stays at Information even with PERF demoted.
- [x] **D2.** Add a narrow `IAppLog` logging seam to `RBLclass.Core` (`SerilogAppLog` forwarding in the AddIn), inject it into `ClassifierService`, and log the previously-silent swallowed failures (per-destination classify failures, safety-copy/attachment-removal/Undo faults, AutoClassify outcome tally).
- [x] **D3.** Fill AddIn/adapter logging gaps: `ShowError` now logs every ribbon/lifecycle exception; the forgotten-attachment and external-recipient send-time guards log when they fire and whether the user proceeded or cancelled; `DeleteItem` and the silent adapter early-exits/S/MIME skips get entry/success/failure breadcrumbs; `Settings.Save` emits one audit line per save summarising the security-relevant settings (never the banner signature).
- [x] Retention investigated: `retainedFileCountLimit: 14` already keeps the log well under 90 days — no change needed.

## v2.9.0.0 — feedback wave: UI polish, classify & banner fixes, attachments (shipped)

> Archived 2026-11-21. B1 is still open: tracked in [ROADMAP](../ROADMAP.md), Open carried-over items [#46].

Planned 2026-07-10 from the open GitHub issues (`/gh-issues` triage). Each
item embeds the implementation decisions made with the user — treat the item
text as the specification.

### A — Pane & search UI

- [x] **A1.** In the folder-search result rows, replace the "open folder"
      affordance's oblique-arrow glyph (top right of the row) with an eye
      glyph (👁), keeping the command, behaviour, placement, and localized
      tooltip exactly as they are. This is a glyph swap only — the eye better
      conveys "look inside this folder" than the arrow. [#45]

- [x] **A2.** Adaptive layout for the pane action buttons (Options, Classify,
      Cancel, Journal): when the task pane is docked horizontally (pane wider
      than tall, e.g. docked above the inbox list), lay these buttons out side
      by side on one line instead of stacked vertically, reclaiming vertical
      space; keep the current stacked layout when the pane is docked
      vertically. Decided: the switch is automatic, driven by the pane's
      aspect ratio (WPF `SizeChanged`/trigger on the view), NOT a user
      setting. [#47]

### B — Bug fixes (banner on reply, Delete key, widening stall)

- [ ] **B1.** Fix external-sender banner stripping on reply/forward. Repro
      (issue #46, reporter's banner is correctly detected — the Settings
      "Diagnose selected mail" reports a match and the filed-copy strip
      works): with the strip-on-reply/forward option ticked, the banner
      remains both in the reply/forward draft and in the sent message.
      Decided: the strip must be applied at draft creation, so the user sees
      the banner disappear before sending. Note: v2.6.0.0 B3 (commit fd2574b)
      already deferred the strip from `NewInspector` to `Inspector.Activate`
      because NewInspector fires before Outlook composes the quoted body —
      verify that deferral actually covers reply/forward drafts, and cover
      inline replies composed in the explorer's reading pane
      (`Explorer.InlineResponse`), which never open an Inspector at all.
      [#46] - shipped, **confirmed still broken on the real target
      (2026-07-28)**: detection works (Settings "Diagnose selected mail"
      still reports a match), but the strip itself still doesn't take
      effect on reply/forward there. Root cause not yet identified — see
      v2.10.0.0 D1, which adds a diagnostic-artifact dump to pin down why
      before attempting a further fix here.

- [x] **B2.** Investigate and fix the dead Delete key after "go to folder"
      (the open-folder action on a search result, which navigates the
      explorer to the target folder): after that navigation, pressing Del no
      longer deletes mail in the target folder; the key works again only
      after the user manually switches to another folder. Known facts: it is
      NOT a plain focus grab by the WPF pane — the maintainer verified that
      even after changing the selection inside the target folder, Del stays
      dead until a manual folder change. Root-cause first (how
      `Explorer.CurrentFolder` is assigned, whether the explorer is activated
      after navigation, whether Outlook's item list ever receives real
      keyboard focus), then fix so Del works immediately after go-to-folder.
      [#49]

- [x] **B3.** Root-cause and fix the ~30 s classify stall tied to
      conversation widening. History: issue #35 reported "Classement en
      cours" for ~30 s with no move on v2.6.0.0; the v2.8.0.0 C2 fix added
      error handling around a conversation-index timeout, but the reporter
      still reproduces on v2.8.0.0 and has now isolated it: classify is
      near-instant with the "process the whole conversation" (widening)
      option unticked and takes ~30 s (the internal ceiling) with it ticked.
      His Inbox/Sent Items are nearly empty and his PSTs live under
      `Documents\FICHIERS OUTLOOK`. **Root cause confirmed from his
      2026-07-07 diagnostic log** (`C:\Users\romai\Downloads\
      rblclass-20260707.log` on the dev machine — reporter's personal data,
      never commit it): every widened classify fails identically in
      `OutlookMailStore.GetConversationSiblings` — the Outlook
      conversation-table enumeration (`TableClass.get_EndOfTable()`) throws
      `COMException` "timeout while obtaining conversation results" (HRESULT
      low word `0x4005`, high word varies) after exactly ~30 s **per item**
      (a 3-item batch stalls 3 × 30 s), with 0 siblings ever collected; the
      v2.8.0.0 C2 handler then logs `widening skipped` and the move itself
      succeeds in ms. So on this profile the conversation-table API never
      works and only burns the 30 s timeout — likely because the machine's
      Windows Search / conversation index is unavailable for those stores.
      Fix directions to evaluate: replace or front-run the conversation
      table with a bounded manual sibling lookup (`Items.Restrict` on
      `ConversationID`/topic over the source folder + candidate destination
      folders, which is how the folder-scoped paths already work), and/or
      after the first timeout in a session disable further widening attempts
      with a user-visible pane notice instead of paying 30 s per item.
      Never a silent 30 s hang. [#35] - pending reporter confirmation (root
      cause fixed and verified in-repo; the reporter's specific broken
      search-index profile needs their own v2.9.0.0 test to confirm)

### C — Sent & meeting items

- [x] **C1.** Add a "copy to inbox" choice to the sent-items arrival
      handling. Today Settings offers moving a newly-arrived sent message to
      the Inbox (or archive / do nothing); add a mode that COPIES it to the
      Inbox, leaving the original in Sent Items. The inbox copy gets the same
      post-sent pin banner as a moved item — including the v2.8.0.0 B1
      "Delete from Inbox" button, which is the escape hatch when the user
      decides not to keep the copy. Rationale (issue #50): classify sent mail
      and keep a copy at hand without toggling the "keep a copy" classify
      option depending on whether the item is sent or received. Needs a new
      settings enum value, a triage-handler branch, and localized labels in
      all three languages (`Strings.resx`/`.fr`/`.de` — ResourceParityTests
      enforces parity). [#50]

- [x] **C2.** Discoverability hint for meeting-invite classify: when the user
      triggers Classify (or Auto-Class) on a selection containing a
      MeetingItem while the opt-in "classify meeting items" setting
      (v2.6.0.0 C1) is OFF, show a one-time informational dialog — "Did you
      know? Classifying meeting invitations is available as an option" —
      naming the exact Settings location, with a "don't show this again"
      checkbox persisted via `ISettingsStore` (settings always persist across
      restarts). The dialog must not interrupt the classify of the other
      selected items — show it after the batch completes. Localized in all
      three languages. [#28] - pending reporter confirmation (implemented
      and built; not exercised interactively this session — reporter asked
      to confirm in v2.9.0.0)

### D — Attachments

- [x] **D1.** Make the saved-attachment location a clickable link. Wherever
      the add-in reports where an attachment (or a saved-as mail file) was
      written — the mail-body annotation added after a save, and any
      pane/status surface showing the path — render the location as a
      hyperlink instead of plain text. Decided: activating the link opens
      Windows Explorer on the containing directory with the file
      pre-selected (`explorer.exe /select,"<full path>"` from WPF surfaces; a
      `file:///` link to the containing folder in the mail-body annotation,
      since an HTML mail body cannot invoke explorer-select), NOT launching
      the file itself. [#51]

- [x] **D2.** Checksum-based duplicate skip on attachment save. Today, saving
      an attachment into a directory that already contains a file with the
      same name writes a "(1)", "(2)"… suffixed copy. New behaviour on every
      attachment save path (classify-time disposition modal, standalone 📎
      modal, save-as): when the target file name already exists, hash the
      existing file and the attachment content (SHA-256, streamed — mind the
      ~1 GB memory budget); if identical, skip the write entirely but still
      record and annotate the attachment as saved exactly like a real save
      (D1's link pointing at the existing file); if the hashes differ, keep
      today's suffix behaviour unchanged. Log each dedupe skip at
      Information level. The compare/decide logic lives in `RBLclass.Core`
      with xUnit coverage. [#52]

## v2.10.0.0 — feedback wave: folder management, pane consolidation, performance, diagnostics (shipped)

> Archived 2026-11-21. D1 is still open: tracked in [ROADMAP](../ROADMAP.md), Open carried-over items [#46].

Planned 2026-07-28 from the open GitHub issues (`/gh-issues` triage). Each
item embeds the implementation decisions made with the user — treat the item
text as the specification.

### A — Folder context menu

- [x] **A1.** Add "Rename folder" to the folder-tree row context menu
      (right-click), as a sibling menu item to the existing "New subfolder"
      (`MainPaneView.xaml:149-154`, `NewSubfolder_Click` at
      `MainPaneView.xaml.cs:241`). Reuse the existing `_promptForName`
      (`Func<string,string>`) prompt-dialog pattern used by
      `CreateSubfolderUnder` (`MainPaneViewModel.cs:1118-1130`), seeded with
      the folder's current name instead of empty. On confirm, call a new
      `IMailStore.RenameFolder(FolderNode folder, string newName)` (declared
      next to `CreateSubfolder` in `IMailStore.cs:236`), implemented in
      `OutlookMailStore.cs` next to `CreateSubfolder` (line 287) using the
      same `ComRef`/`GetFolderFromID` pattern — set `Folder.Name` on the live
      Outlook folder, catching `COMException` (duplicate name, or Outlook
      refusing a special/reserved folder) and reporting it via the existing
      status-line pattern rather than throwing. On success, reindex via
      `FolderIndexService.ReindexStore(folder.StoreId)`, mirroring the
      "update folders index" step already done for create-subfolder
      (`RblClassAddIn.cs:205-219`). Decided: disable/hide "Rename" when the
      target is a store's root folder (PST root) — renaming a store root is
      unusual and can confuse Outlook's store list; any other
      folder/subfolder is renamable. New resource key + FR/DE translation
      for the menu label and any new status strings (`ResourceParityTests`).
      Cover the root-rename-blocking check with an xUnit test in
      `RBLclass.Core` (the COM rename itself is untestable off a live
      Outlook). [#54]

- [x] **A2.** Add "Delete subfolder" to the same context menu, as the
      inverse of "New subfolder". Add `IMailStore.DeleteFolder(FolderNode
      folder)` next to `RenameFolder`/`CreateSubfolder` in `IMailStore.cs`
      and `OutlookMailStore.cs`, calling Outlook's `Folder.Delete()`
      (default semantics: moves the folder to Deleted Items unless it's
      already there). Decided: before deleting, show a confirmation dialog
      reusing the existing confirm-dialog pattern (`Func<int, bool?>`, same
      shape as `_confirmMarkTasksComplete`, `MainPaneViewModel.cs:32`) whose
      text states how many mail items and subfolders will be removed; if the
      user confirms, proceed with the delete even when the folder is
      non-empty — the count shown in the confirmation is the safety
      mechanism, there is no separate hard block on non-empty folders.
      Reindex the store via `FolderIndexService.ReindexStore` on success,
      exactly as A1. Disable/hide "Delete" for a store's root folder (same
      restriction as A1). New resource keys (menu label, confirmation
      dialog text with a `{n}` item-count placeholder using the
      `_One`/`_Other` plural pattern) translated in FR/DE, checked by
      `ResourceParityTests`. [#57]

### B — Pane UI

- [x] **B1.** Move the "Settings" and "Update folders" actions from the
      RBLclass ribbon (`Ribbon.xml:15-28`, `SettingsButton` /
      `RefreshFoldersButton`, `OnSettingsClick` / `OnRefreshFoldersClick`)
      into the RBLclass pane, removing them from the ribbon (`TogglePaneButton`
      and `AboutButton` stay — only these two move). Decided: they join the
      existing icon-button row at the top of the pane
      (`MainPaneView.xaml:80-119`, the `DockPanel LastChildFill="False"`
      currently holding AutoClass / SaveAs / ManageAttachments and the
      index-status `Ellipse`), as two more docked-right icon buttons using
      the same `FontFamily="Segoe UI Emoji"` styling as SaveAs /
      ManageAttachments. Settings renders as a gearbox emoji (⚙️),
      Update-folders as a circle-arrow emoji (🔄); clicking them invokes the
      same logic currently wired to `OnSettingsClick` / `OnRefreshFoldersClick`
      (extract to shared methods callable from both the new pane
      code-behind Click handlers and the ribbon callbacks — the ribbon
      buttons are being removed, so these become the only call site). New
      tooltip resource keys mirroring the ribbon's existing
      `screentip`/`supertip` text, translated in FR/DE (`ResourceParityTests`
      already covers `Ribbon.xml`/`.fr.xml`/`.de.xml` parity — verify the new
      pane-side keys are covered too). [#55]

### C — Performance

- [x] **C1.** Fix the long delay in the pane's "N selected" indicator on
      large Outlook selections. Root cause: `OnExplorerSelectionChange`
      (`RblClassAddIn.cs:720-746`) synchronously re-walks the full COM
      `Selection` via `GetSelectedItems()` (`OutlookMailStore.cs:219-274`,
      one per-item COM marshalling call per selected item) on every
      `SelectionChange` event, on the STA/UI thread, with no throttling —
      for large selections (hundreds to thousands of items) this blocks
      Outlook's UI without feedback. Decided fix: debounce only, no hard cap
      on the number of items processed — apply the existing
      `SearchDebounceMs`-style coalescing already used for folder search
      (`Settings.cs` / `SettingsKeys.cs:43`, default 200 ms, clamped to
      `MaxSearchDebounceMs` = 2000 ms) to `Explorer.SelectionChange` events,
      so a burst of rapid selection changes (e.g. shift-click-dragging
      across thousands of rows) collapses into a single walk after the user
      pauses, instead of one full walk per intermediate event. The debounce
      timer must run off the STA thread's own dispatcher (not a
      thread-pool timer calling back into COM), per the add-in's threading
      rule. Add xUnit coverage in `RBLclass.Core` for the debounce-decision
      logic (a small class deciding "have enough events settled to fire"
      given an injectable clock), and manually verify by selecting large
      ranges (1000+ items) in the dev Outlook and confirming the pane no
      longer stalls the UI. [#56]

### D — Banner diagnostics

- [ ] **D1.** Add a diagnostic-artifact dump to the existing "Diagnose
      selected mail" Settings action (`SettingsViewModel.DiagnoseSelectedMail`,
      `SettingsViewModel.cs:482-513`), to unblock root-causing why v2.9.0.0
      B1's reply/forward banner strip (issue #46) detects the banner
      correctly but still fails to strip it on the real 32-bit target.
      Today `ExternalBannerStripper.Diagnose` (`ExternalBannerStripper.cs:135`)
      returns only `BannerDiagnosticOutcome` / `SignatureLength`, no
      substring or position — extend it (or add a sibling method) to also
      surface the matched range: the ordinal index (`Strip`'s exact-match
      path, `ExternalBannerStripper.cs:143-145`) or the regex `Match.Index` /
      `Length` (tolerant path, lines 102-118), so a caller can slice the
      original body. Decided: add a manual "Save diagnostic" action next to
      the existing Diagnose button/result in Settings — on-demand only, not
      auto-captured on every live strip attempt — that writes a timestamped
      `.html` file to a new `%LocalAppData%\RBLclass\diagnostics\` folder
      (mirroring the existing `logs` / SQLite-under-`%LocalAppData%\RBLclass\`
      convention, `RblClassAddIn.cs:1300-1328`). The file contains the
      complete original HTML body with the matched banner substring (if any)
      wrapped in a visibly distinct inline style (e.g. a red dashed border),
      plus a small metadata comment block (match outcome, exact-vs-tolerant,
      signature length, timestamp), so the reporter can send back one
      self-contained file showing exactly what would/wouldn't be stripped.
      This persists mail body content to disk for the first time outside
      the SQLite index — flag it in the security-posture doc pass
      (`docs/security.md` + `website/docs/security.md`, via the
      `docs-writer` subagent): new data written to disk, on-demand and
      user-triggered only, local machine, same trust boundary as the
      existing log files. New resource keys for the "Save diagnostic"
      button/tooltip and a save-confirmation status message, translated
      FR/DE. [#46] - shipped, pending tests (not verifiable on the dev
      machine per the testing strategy above; no live transport-rule
      banner is injected into mail on this dev account, so the "Save
      diagnostic file" action, its highlighted-match rendering, and the
      lazy `diagnostics\` folder creation need a real-target re-check
      before pilot)

## v2.11.0.0 — feedback wave: install & encryption diagnostics, keyboard flow, attachments (shipped)

> Archived 2026-11-21. A3 and C2 are still open: tracked in [ROADMAP](../ROADMAP.md), Open carried-over items [#62].

Planned 2026-11-05 from the open GitHub issues (`/gh-issues` triage). Each
item embeds the implementation decisions made with the user — treat the item
text as the specification.

### A — Fixes & diagnostics

- [x] **A1.** Fix the saved-attachment folder link for paths containing
      accented characters. `AttachmentLabelFormatter.PathToFileUri`
      (`AttachmentLabelFormatter.cs:68`) uses `Uri.AbsoluteUri`, which
      percent-encodes non-ASCII as UTF-8 (`é` → `%C3%A9`); the likely fault
      is the Windows shell decoding such `file:` URLs with the local code
      page. Decided: keep Unicode characters raw in the href and percent-encode
      only spaces and URI-reserved characters (`%`, `#`, `?`), still
      HTML-encoded for the attribute. Add xUnit cases with accented paths,
      spaces and a UNC path (existing tests are ASCII-only). The in-modal link
      goes through `Process.Start("explorer.exe", ...)` and is not affected.
      Confirm the root cause by clicking a generated link on Windows. [#64]

- [x] **A2.** Make MSI failures diagnosable and correct the install guide.
      Set the `MsiLogging` property in `installer/Package.wxs` so every
      install writes a verbose log to `%TEMP%` with no command-line switch.
      Add an install-troubleshooting section to
      `website/docs/user-guide/installation.md`: the log location, the
      `msiexec /i <msi> /l*vx <log>` command, the `MsiInstaller` entries in
      Event Viewer, and the PowerShell install kit as a fallback. Fix that
      page's install steps: add "close Outlook first" (today only
      `docs/installation-guide.html` says so) and drop "Next → Install →
      Finish" (the package has no `WixUI` wizard). Run the discriminating
      test on Windows and record the result on the issue: install 2.9.1.0,
      leave Outlook open, run the 2.10.0.0 MSI with `/l*vx`. The actual fix
      is the carried-over item above. [#66] - pending reporter confirmation
      (#66 stays open for the carried-over fix)

- [ ] **A3.** Harden and instrument encrypted-mail handling during classify.
      Analysis (2026-11-05): the default single-destination classify is a pure
      `mail.Move()` and cannot remove encryption from the code's side, but
      two gaps exist. (1) `MarkTaskComplete` (`OutlookMailStore.cs:1197`,
      reached from classify) and `MarkTaskIncomplete` (`:820`) call
      `mail.Save()` with no S/MIME guard. (2) Every other guard relies only on
      `MessageClass` starting with `IPM.Note.SMIME`; if Stormshield exposes
      another class, none fires and even mark-as-read saves the item.
      Decided: one shared protected-mail check in the adapter
      (`MessageClass` plus `PR_SECURITY_FLAGS` via `PropertyAccessor` —
      validate that property on the target first) replacing the scattered
      `StartsWith` checks; guard both task-flag saves; add a trace gated by
      `DiagnosticLogging` recording the class and security flags before and
      after classify, plus which path ran (move, copy, safety copy). No
      re-encryption and no user prompt in this item: decide those only if the
      trace implicates RBLclass. Update `docs/security.md` and
      `website/docs/security.md` through `docs-writer`. Not verifiable on
      the dev machine per the testing strategy above; re-check on the real
      target. [#62] - shipped, pending tests (not verifiable on the dev machine
      per the testing strategy above). Decided 2026-11-05: PR_SECURITY_FLAGS is
      log only, protection stays class-based until validated. Re-verify on the
      real target before pilot: with Diagnostic logging on, classify one
      protected mail per path (move, copy, safety copy, keep a copy) and read
      the real class and flags from the `Protection trace` lines, then decide
      whether the flags or another class prefix should drive the check. On
      plain mail the dev machine shows `flags=0x00000000`, so the property is
      present there. Also pending a live regression pass on ordinary mail:
      copy, safety copy, mark as read, attachment removal, conversation
      widening (only move, task-flag completion and Undo were exercised).
      Target result 2026-11-13 (v2.11.0.0, Stormshield): with the smart card
      inserted, "remove attachments" silently stripped an encrypted mail's
      attachments and body images (`encryptedKept=0`). Stormshield decrypts the
      opened item in memory and it reports `IPM.Note`, so the class-only check
      passed. Fixed on develop 2026-11-14: `ProtectedMailPolicy.Evaluate`
      also treats as protected a protected **stored** class (parent folder
      contents table, narrowed by Internet message id, matched with
      `CompareEntryIDs`), a `smime.p7m`/`smime.p7s` part, and unsaved changes
      on open (not for drafts). A non-class detection logs `protected mail
      detected by <Reason>` at Information. Dev machine: verified by driving
      the adapter (plain mail still stripped, simulated in-memory decryption
      and dirty-on-open refused). Without the card, Classify did nothing and
      logged nothing; unreadable selected items and empty classifies are now
      logged (Warning / Information). Re-verify on the target: card in,
      classify an encrypted mail with "remove attachments" on: the attachments
      stay, and the log shows `protected mail detected by StoredMessageClass`
      (or another reason). Card out: send the new `GetSelectedItems` warning
      line.
      E2E on the dev machine 2026-11-14 (fake mails, add-in driven through UI
      Automation): Outlook's `Move`/`Copy` **write the in-memory state** of an
      item held modified by another handle (`Close(olDiscard)` does not stop
      it). A Stormshield-decrypted mail is therefore filed decrypted, which
      likely explains the original "loses its encryption" report and cannot be
      prevented through the object model. Fix: `ClassifierService` decides
      protection once on the source **before** filing and makes no write on
      any filed copy of a protected source (no strip, body images, label,
      banner, mark as read, task flag). It always logs a Warning `Protected
      mail filed by … no longer reads as protected` when the filed item came
      out plain. E2E green: plain move, strip via the window, mark as read,
      manage attachments, Undo, empty selection, and the Stormshield simulation
      (attachments kept, not marked read, `encryptedKept=1`, warning logged).
      Target: with the card in, check that the attachments stay and look for
      that warning (it means the mail was filed decrypted).
### B — Pane & keyboard

- [x] **B1.** Remove the selection-count label at the bottom of the pane:
      the row 8 `TextBlock` in `MainPaneView.xaml`, `SelectionSummary`, and
      the `Status_MailSelected_One`/`_Other` and
      `Status_NoMailSelectedSummary` keys in all three `.resx` files. Keep
      `ClearPin()` firing on a selection change, and keep the banner
      availability refresh in `RefreshSelection`. Drop the
      `GetSelectedItems().Count` COM call from the debounced selection tick
      if nothing else needs it. [#58]

- [x] **B2.** Expose the online documentation
      (`https://reboulip.github.io/rblclass/`, English only) from a link in
      the About window and a "?" button in the pane header next to Settings,
      opened in the default browser. New label/tooltip resource keys,
      translated FR/DE. [#59]

- [x] **B3.** Return keyboard focus to the search box after pane actions and
      add a `/` shortcut. Replace `FocusSearchAfterClassify` with one shared
      path that covers every pane action and modal close, and that does what
      `FocusSearch` does (`Focus`, `Keyboard.Focus`, `SelectAll`); if WPF
      focus alone does not take because Outlook holds Win32 focus, also move it
      to the pane's host window. Decided: "open folder" is excluded, focus
      stays in the explorer so the Delete key keeps working (v2.9.0.0 B2,
      issue #49). `/`: when the search box does not have focus it takes
      focus and the character is swallowed; when it has focus, `/` is typed
      normally. Implement through `PreviewTextInput`, not `KeyDown`, so it
      works on AZERTY and the numeric keypad. Pane-scoped only: no global
      keyboard hook (EDR risk); from Outlook the existing ribbon keytip reaches
      the pane. Update the walkthrough shortcuts table. [#60]

- [x] **B4.** Make keyboard focus more visible, moderately: from subtle to a
      little less subtle, not vivid. `Controls.xaml` has no keyboard-focus
      state today. Add a `Theme.FocusBorder` brush to both
      `Brushes.Light.xaml` and `Brushes.Dark.xaml`, and a thin accent-coloured
      outline on buttons, check boxes, radio buttons, text boxes, combo boxes
      and list items, driven by `IsKeyboardFocused`. Validate the contrast in
      dark mode with the user. [#61]

### C — Attachments

- [x] **C1.** Add a default attachment folder and a grouped view to the
      attachment disposition modal. Setting: a new key with a Browse control in
      Settings → Attachments, pre-filling the target directory; ignored
      (fields left empty) when the folder is unreachable, e.g. a disconnected
      share. Grouped view: when `UseSameDirectory` is ticked, hide the per-row
      controls and show one set of options (delete / save to / keep, plus the
      path) and a list of the affected attachments with their sizes, updating
      each row behind the scenes. That box is ticked by default, so the grouped
      view becomes the standard one, and it now covers the action as well as
      the folder: reword the label in EN/FR/DE. Unticking restores the per-row
      view. Keep the apply-to-all rule testable in `RBLclass.Core` with xUnit
      cases, and preserve the v2.7.1.0 A2 rule when returning to per-row mode.
      [#65] Decided 2026-11-07 (live review): the tick box became three modes,
      default first: same action and folder for all attachments (the grouped
      view); an action per attachment with one folder for the saved ones; an
      action and a folder per attachment. The "Delete all" button is removed,
      covered by the first mode.

- [ ] **C2.** Add an opt-in option to treat body images as removable
      attachments (requires: C1, A3). New setting in Settings → Attachments,
      default off; v2.5.0.0 B2's exclusion of inline images stays the default.
      When on, `cid:` inline images appear in the disposition modal, marked
      as such and defaulting to Keep; embedded items, OLE objects and remote
      `http` images stay out of scope. Silent removal
      (`DeleteSilently`) never removes them, since signature logos are inline
      images and the removal is irreversible. Removing an image must also
      rewrite the HTML: add a pure `RBLclass.Core` rewriter, with xUnit tests,
      replacing the matching `<img src="cid:...">` by a short text placeholder
      so no broken image remains; expose the content id on `AttachmentInfo`.
      Encrypted mail stays untouched. New resource keys translated FR/DE. [#63]
      Decided 2026-11-07: with the option on, body images are offered even when
      "remove attachments" is off (only the images are then listed); the
      placeholder reads "[Image removed]" / "[Image supprimée]" / "[Bild
      entfernt]"; Save to also applies to body images. Verified live on the dev
      machine: single and partial image removal, classify with removal off. -
      shipped, pending tests (not verifiable on the dev machine per the testing
      strategy above; re-verify on the real target that an S/MIME/Stormshield
      mail with body images keeps them and its body untouched before pilot)

## v2.12.0.0 — experimental: re-encrypted filing of protected mail (shipped)

> Archived 2026-11-21. A8 is still open: tracked in [ROADMAP](../ROADMAP.md), Open carried-over items [#62].

Context (GH #62, 2026-11-13/14). With the Stormshield smart card inserted, an
encrypted mail is decrypted **in memory** when Outlook opens it, and Outlook's
`MailItem.Move`/`Copy` write that in-memory (decrypted) state into the filed
item. That plaintext write cannot be prevented through the object model; it was
reproduced on the dev machine by simulation. The standard path (develop,
commits 981f1fe and d50c145) never writes to a protected mail's filed copy.
That stops the attachment loss, but the copy can still be filed decrypted (a
Warning is logged when it is). This milestone adds an **opt-in experimental
alternative**: rebuild the protected mail from its decrypted in-memory view,
apply the requested actions to that content, re-encrypt it as standard S/MIME
to the user's certificate, and create the filed item directly as
`IPM.Note.SMIME`, so that no plaintext is ever stored. The standard path stays
the default and remains the fallback whenever re-encryption cannot be done.

Spike (2026-11-14, dev machine, PowerShell): a MIME entity encrypted with
`System.Security.Cryptography.Pkcs.EnvelopedCms` (`System.Security.dll`, part
of .NET Framework 4.8, so no new package) to a self-signed certificate in
`CurrentUser\My` (AES-256-CBC; key transport defaults to RSA-OAEP), saved as a
`smime.p7m` attachment (`PR_ATTACH_MIME_TAG` = `application/pkcs7-mime`) on an
item created with the PostItem trick (`Folder.Items.Add(olPostItem)`,
attachment, Save, then `MessageClass = "IPM.Note.SMIME"`, Save). Outlook's own
S/MIME layer decrypts it: body and kept attachment are shown, and the list
shows the encrypted icon. `MailItem.SaveAs(..., olMSG)` writes the stored
(encrypted) form. Script: `tools/Inspect-SmimeMsg.ps1` reads a `.msg` offline
and reports the CMS envelope (content cipher, key transport, recipient
certificate, whether it is in `CurrentUser\My`). **Step 0, pending:** its
output for a real Stormshield mail on the target decides the key-transport
default in A1 and confirms Stormshield uses standard CMS. If it does not, the
feature is shelved.

Decisions (user, 2026-11-14): experimental and off by default, in a new
"Experimental" Settings tab. When on, **every** protected mail that is readable
(decrypted in memory) is rebuilt, even with no action requested. The sender's
signature on the filed copy is lost (accepted). The original goes to Deleted
Items (recoverable); if that move writes it decrypted, it is permanently
deleted and logged. Certificate: automatic, with an explicit choice in
Settings. **No Undo** for re-encrypted filings (accepted).

Not verifiable on the dev machine: Stormshield itself. Everything else is
verifiable with a self-signed certificate and the simulated-Stormshield
technique (memory `e2e-outlook-testing`: stored `IPM.Note.SMIME`, class
rewritten to `IPM.Note` in memory by a held COM reference). Target check: a
pilot machine opens a re-encrypted mail with Stormshield.

### A — Re-encrypted filing (experimental)

- [x] **A1.** Settings: an "Experimental" tab and the re-encryption settings.
      Add a new `TabItem` at the end of `SettingsWindow.xaml` (after
      General), header key `Settings_Nav_Experimental`, holding one
      `GroupBox` `Settings_Experimental_Header`. Start it with a warning
      `TextBlock` (`Settings_Experimental_Warning`) saying that these features
      are experimental and meant for pilot testing. Then:
      (1) CheckBox `Settings_ReencryptProtected_CheckBox` "Re-encrypt
      protected mail when filing" with an info tooltip
      (`Settings_ReencryptProtected_Info`) explaining that the mail is
      rebuilt and re-encrypted to your certificate, the sender's signature is
      not kept, the filing cannot be undone, the original goes to Deleted
      Items, and it only applies while the mail is readable (smart card
      inserted);
      (2) a ComboBox `Settings_ReencryptCertificate_Label` listing "Automatic"
      plus every candidate certificate (see A4 `ListEncryptionCertificates`)
      shown as `Subject (expires yyyy-MM-dd)`; the selected value is stored
      as a thumbprint, empty = Automatic;
      (3) a read-only line showing the certificate Automatic currently
      resolves to, or `Settings_ReencryptCertificate_None` when none.
      Keys in `SettingsKeys`: `ReencryptProtectedOnFile` (bool, default
      false), `ReencryptCertificateThumbprint` (string, default empty),
      `ReencryptKeyTransport` (string `Oaep` | `Pkcs1`, default from step 0,
      `Oaep` until then; no UI). Add them to `Settings` (load/save pairs, like
      `TreatInlineImagesAsAttachments`) and to `SettingsViewModel`
      (`Apply(...)` pattern, persisted immediately per memory
      `settings-always-persistent`). The certificate list comes from a
      `TaskPaneServices` delegate wired in `RblClassAddIn` to the adapter.
      All strings go in EN/FR/DE resx (`ResourceParityTests`). Update
      `SettingsTests` for the new keys.

- [x] **A2.** Core: `MimeEntityBuilder`, a pure MIME writer
      (`RBLclass.Core`). Input: a new `MailContent` model (`HtmlBody`,
      `PlainBody` (either may be null), `IReadOnlyList<MailPart> Parts`, where
      `MailPart` = `FileName`, `ContentType`, `byte[] Data`, `ContentId`
      (null when not inline), `IsInline`). Output: the bytes of one MIME
      entity (no RFC 822 envelope headers: those stay outside the encryption,
      as S/MIME specifies). Structure: `multipart/mixed` when there are
      non-inline parts, wrapping `multipart/related` (HTML body plus inline
      parts with `Content-ID: <cid>`) when there are inline parts, wrapping
      `multipart/alternative` when both bodies exist; a single part collapses
      to itself. Bodies `charset=utf-8`, `Content-Transfer-Encoding:
      quoted-printable` or base64. Parts are base64 with 76-character lines.
      CRLF line ends everywhere. Boundaries are random and must not occur in
      the content. Non-ASCII file names use RFC 2231 `filename*=utf-8''...`
      plus an RFC 2047 encoded `name=` fallback. A missing `ContentType`
      defaults to `application/octet-stream`. xUnit tests (new file
      `MimeEntityBuilderTests`): each structure shape, a boundary collision
      guard, line length ≤ 998 and base64 ≤ 76, round-trip of a non-ASCII
      file name, an empty body, a large part (5 MB) with no excessive
      allocation (stream to a `MemoryStream`).

- [x] **A3.** Core: `ReencryptContentPlanner`, which applies the classify
      actions to a decrypted `MailContent` and returns the content to
      encrypt, with no Outlook. Inputs: the `MailContent`, the
      `ClassifyRequest` options for that item (`RemoveAttachments`, the
      item's `AttachmentDisposition` rows, `BannerSignature`, `LabelOptions`).
      Rules, the same as the standard path's on plain mail:
      strip-all removes every non-inline part. Dispositions remove parts by
      attachment id: map ids to part order, documented in the adapter (A5).
      "Save to" rows are saved by the caller before planning (A6). Inline
      image rows rewrite `HtmlBody` through `InlineImageHtmlRewriter` and drop
      the part. The banner goes through `ExternalBannerStripper.Strip`. The
      label goes through `AttachmentLabelFormatter` and is inserted after
      `<body>` (reuse the adapter's `InsertAfterBodyTag` logic, moved to Core
      as a public static helper). Output: the new `MailContent` plus the
      labelled rows. xUnit tests mirroring `ClassifierServiceInlineImageTests`
      and the strip/keep/label cases.

- [x] **A4.** Core interface `IMailRebuilder`, implemented by the adapter,
      kept separate from `IMailStore` so the shipped contract does not grow
      (memory `prefer-isolated-new-ui-over-retrofit`):
      `MailContent ReadDecryptedContent(MailItemRef item)` returns null when
      the item is not readable decrypted, i.e. it still shows its
      `smime.p7m` part, or the read failed.
      `byte[] Encrypt(byte[] mimeEntity, string certificateThumbprint, string keyTransport)`
      returns null on failure, after logging the reason.
      `MailItemRef CreateEncryptedCopy(MailItemRef source, FolderNode destination, byte[] smimePayload, bool unread)`
      copies the envelope from `source`.
      `bool VerifyEncryptedCopy(MailItemRef copy, byte[] smimePayload)`.
      `CertificateInfo ResolveCertificate(string preferredThumbprint)` returns
      null when nothing is usable.
      `IReadOnlyList<CertificateInfo> ListEncryptionCertificates()`.
      `CertificateInfo` = Thumbprint, Subject, NotAfter, HasPrivateKey.
      Also define `ReencryptOutcome` (see A6).

- [x] **A5.** Adapter: `OutlookMailRebuilder : IMailRebuilder`
      (`RBLclass.Outlook.Adapter`; COM rules from CLAUDE.md: `ComRef`,
      indexed loops, no chained property access, STA only).
      *ReadDecryptedContent*: resolve the item and refuse (return null) when
      `ProtectedMailPolicy.IsSmimePayloadName` matches any attachment (not
      decrypted: the card is absent). Read `HTMLBody` when `BodyFormat` is
      HTML, else `Body`. For each attachment (by index), read the bytes from
      `PR_ATTACH_DATA_BIN` (0x37010102) through the attachment's
      `PropertyAccessor`. If that throws (large attachments can), fall back to
      `SaveAsFile` into `%LocalAppData%\RBLclass\tmp\<guid>`, read the bytes,
      overwrite the file with zeros, delete it, and log the fallback (no file
      name). Also read `PR_ATTACH_MIME_TAG` and the content id
      (0x3712001F), mark inline with the same rule as `IsInlineAttachment`,
      and skip embedded items and OLE parts (`olEmbeddeditem`/`olOLE`), which
      cannot be faithfully rebuilt: refuse the whole item (return null) when
      one is present, so it falls back to the standard path. Cap the total at
      50 MB (refuse above it; 32-bit memory budget). Part order = attachment
      index order, and that is the id mapping A3 uses.
      *Encrypt*: `EnvelopedCms` with
      `ContentInfo(new Oid("1.2.840.113549.1.7.1"), mime)`, AES-256-CBC
      (`2.16.840.1.101.3.4.1.42`), and `CmsRecipient` with the certificate.
      For `Pkcs1`, use the `CmsRecipient(cert, RSAEncryptionPadding.Pkcs1)`
      overload (check it exists in .NET 4.8; if not, record that
      `Pkcs1` is unsupported and use the default). Encrypt to the user's
      certificate only.
      *CreateEncryptedCopy*: write `smimePayload` to a temp file (it is
      ciphertext, so that is safe), then in `destination` create
      `Items.Add(olPostItem)`. Set `Subject` from the source and leave the
      body empty (no plaintext outside the envelope). Add the attachment
      `smime.p7m`, set `PR_ATTACH_MIME_TAG` = `application/pkcs7-mime`, Save,
      set `MessageClass = "IPM.Note"`, Save, release, and re-open as a
      `MailItem`. Add the source's recipients by type (`Recipients.Add`
      with the SMTP address, `Type` To/Cc/Bcc). Copy through
      `PropertyAccessor`: sender and sent-representing name, address and
      address type (0x0C1A001F, 0x0C1F001F, 0x0C1E001F, 0x0042001F,
      0x0065001F, 0x0064001F), `PR_CLIENT_SUBMIT_TIME` (0x00390040),
      `PR_MESSAGE_DELIVERY_TIME` (0x0E060040), `PR_INTERNET_MESSAGE_ID`
      (0x1035001F), conversation topic and index (0x0070001F, 0x00710102),
      importance and categories. Set `UnRead`, Save, set
      `MessageClass = "IPM.Note.SMIME"`, Save, release, delete the temp file.
      A property Outlook refuses is logged (tag only) and skipped. Verify on
      dev which ones stick, and record the result in this item. Return the
      new item's ref.
      *VerifyEncryptedCopy*: re-open fresh and check, through
      `ReadStoredProtectedClass` (move it to a shared internal helper), that
      the stored class is `IPM.Note.SMIME` and that the folder table holds
      the row. Without decrypting, compare the stored `smime.p7m` bytes
      with `smimePayload` (read via `PR_ATTACH_DATA_BIN`). If Outlook's
      S/MIME layer hides the attachment, accept the class check alone and log
      that.
      *ResolveCertificate / ListEncryptionCertificates*: certificates in
      `CurrentUser\My` that are valid now, have an RSA key, have
      `KeyEncipherment` in Key Usage (or no Key Usage extension), and have
      `1.3.6.1.5.5.7.3.4` in EKU (or no EKU). Automatic = among those, the
      ones whose subject e-mail or SAN rfc822 equals the SMTP address of the
      default store's account (`Session.Accounts`, matching store, else the
      first account); then the latest `NotAfter`. An explicit thumbprint that
      is no longer valid resolves to null (no silent switch).

- [x] **A6.** `ClassifierService`: the re-encrypted path
      (requires: A2, A3, A4). `ClassifierService` takes an optional
      `IMailRebuilder` (ctor parameter, default null = feature unavailable).
      `ClassifyOptions` gains `ReencryptProtected` (bool) and
      `ReencryptCertificateThumbprint`; the AddIn sets them from A1's
      settings for both Classify and AutoClass. In `ProcessItem`, when
      `wasProtected && ReencryptProtected && rebuilder != null`:
      1. Run the "Save to" dispositions on the source first
         (`SaveAttachmentToFile` on the source, which is readable in memory).
      2. `ReadDecryptedContent`. If it returns null, use the standard path.
      3. Plan (A3), then `MimeEntityBuilder`, then `ResolveCertificate`, then
         `Encrypt`. If any step fails, use the standard path.
      4. For every destination, `CreateEncryptedCopy` (unread = the source's
         unread state, or false when MarkAsRead is requested) followed by
         `VerifyEncryptedCopy`. If a copy fails verification, delete that
         copy, use the standard path for this item, and stop.
      5. Unless KeepCopy is on: move the source to its store's Deleted Items
         (`GetDeletedItemsFolder` + `MoveItemToFolder`, markAsRead false).
         Check the moved item with `IsEncryptedMail`; when it no longer reads
         as protected (plaintext written by Move), delete it permanently
         (`Delete()` in Deleted Items, as `DeleteItemPermanently` does) and
         log a Warning.
      6. Count the item as processed and filed. Record history (conversation
         key read beforehand) but add nothing to the Undo plan.
      `ClassifyResult` gains `Reencrypted` and `ReencryptFallbacks` counts.
      The AddIn status appends `Status_Classify_Reencrypted_One/_Other`
      ("{0} encrypted mail(s) re-encrypted (cannot be undone)") and
      `Status_Classify_ReencryptFallback_One/_Other` ("{0} encrypted mail(s)
      filed without re-encryption — see the log"), plus a one-time-per-session
      `Status_Classify_ReencryptNoCertificate` when no certificate resolves.
      Logs at Information: "Re-encrypted filing: {Destinations}
      destination(s), {Parts} part(s), cert {Thumbprint}" (no subject, no file
      names). Fallback reasons at Warning. The attachment window: in
      re-encryption mode, protected items readable in memory are offered like
      plain mail (they will be rebuilt). Gate this in
      `AllAttachmentBearersAreEncrypted`/`GatherAttachments` on the setting.
      xUnit tests with NSubstitute for `IMailRebuilder`: the happy path
      (copies, source to Deleted Items, no undo entry, counts), each fallback
      point, KeepCopy keeps the source, a plaintext Deleted Items copy is
      deleted permanently, the setting off means the standard path is
      unchanged.
      Done (2026-11-14): eligibility in `ReencryptEligibility` (new outcome
      `NotEncrypted` for mail protected only by unsaved changes); the original
      already in Deleted Items is not moved; a regular part keeps its
      `Content-ID`, so an unflagged body image (IMAP/PST stores) still renders
      in the rebuilt copy. Verified on dev with the simulated-Stormshield
      technique: move + strip + modal + mark as read, keep a copy, no strip
      (copy opens decrypted, label, accents, image, original hard-deleted after
      the plaintext move). Stormshield itself is re-checked in A8.

- [x] **A7.** Docs and the security posture (through `doc-keeper`):
      `website/docs/user-guide/walkthrough.md` (the Experimental tab, what it
      does, the limits: no undo, signature lost, card required),
      `docs/security.md` + `website/docs/security.md` (a new trust boundary:
      the add-in handles decrypted content in memory and performs S/MIME
      encryption; the temp-file fallback is zeroed then deleted; the
      certificate is public-key only; the original in Deleted Items, or
      permanently deleted when written decrypted), and
      `developer-guide/architecture.md` (`IMailRebuilder`, the MIME builder,
      use of in-box `System.Security.dll`). Commit
      `tools/Inspect-SmimeMsg.ps1` and mention it in
      `developer-guide/testing.md`.

- [ ] **A8.** E2E verification on the dev machine (no user involvement; memory
      `e2e-outlook-testing`). Use the self-signed `CN=RBLclass E2E Test`
      certificate (created by the spike) and fake mails made truly encrypted
      by the spike technique, plus the in-memory-decrypted simulation for the
      Stormshield case. Scenarios: re-encrypted move with strip, banner and
      mark as read; KeepCopy; two destinations; card-absent (payload visible),
      which falls back to the standard path; no certificate, which falls back
      with the status line; an embedded `.msg` attachment, which falls back;
      the setting off, which runs the standard path. Each re-encrypted copy
      must open decrypted in Outlook with the expected parts. Target: a pilot
      machine with Stormshield opens a re-encrypted mail. [#62]
      Dev results (2026-11-14, simulated Stormshield, explicit E2E
      certificate): move + strip + attachment window + mark as read, keep a
      copy, no strip, two destinations all re-encrypted; each copy is stored
      `IPM.Note.SMIME`, opens decrypted with the expected parts, accents and
      label, unread state kept or cleared as asked; the original written
      decrypted by the move to Deleted Items is hard-deleted (logged). Card
      absent (payload visible) takes the standard path silently
      (`NotDecrypted`, fixed here: the eligibility now reads the parts when
      the live class is protected); embedded `.msg` falls back (`Unsupported`);
      Automatic with no usable certificate falls back with both status lines;
      setting off runs the standard path. Certificate pick in Settings saved.
      No temp file left. - shipped, pending tests (not verifiable on the dev
      machine per the testing strategy above; re-verify on the real target
      before pilot: with the card inserted, a classified Stormshield mail is
      filed re-encrypted and opens decrypted in Outlook with all its parts;
      Stormshield accepts the RSA-OAEP key transport (step 0); the external
      banner is stripped from the rebuilt body; card absent files normally)

## Carried-over items closed (archived 2026-11-21)

Moved out of ROADMAP.md's "Open carried-over items" on 2026-11-21, together
with the completed Phase 9 hardening items.

### Resolved

- [x] **Task pane rendered partly behind the navigation rail when docked
      top/bottom.** RESOLVED 2026-07-11 (plus a hard-freeze regression fixed the
      same day — see resolution block at the end of this item).
      Discovered during v2.9.0.0 A2 verification: when the CTP's
      `DockPosition` is Top or Bottom, its left edge renders underneath
      Outlook's left navigation icon rail (Mail/Calendar/People/Apps), both on
      live redock and on a fresh Outlook restart with the position persisted
      from a prior session — not a one-off repaint glitch. Left/Right docking
      (the common case, used throughout normal dev/testing) is unaffected.
      **A real, separate bug found and fixed along the way:** `EnsureTaskPane`
      unconditionally set `_taskPane.Width = 360` regardless of dock position;
      Office's `CustomTaskPane` only accepts `Width` while docked Left/Right
      and only `Height` while docked Top/Bottom, so setting the wrong one
      throws `COMException 0x80004005` — this crashed `OnStartupComplete`
      whenever the persisted dock position was Top/Bottom. Fixed by
      `ApplyDockSize` (`RblClassAddIn.cs`), which sets `Width` or `Height`
      based on the pane's actual `DockPosition`; also invoked on every
      `DockPositionStateChange` so live redocking gets correctly resized too.
      **Ruled out for the remaining visual overlap** (crash fix alone did not
      resolve it): deferring the initial `EnsureTaskPane()` + `Visible = true`
      call to `DispatcherPriority.ApplicationIdle` in `OnStartupComplete`, on
      the theory that showing the pane raced the Explorer's own nav-rail
      layout — reverted, no effect, overlap reproduces identically on a cold
      restart either way. Likely a genuine Office/Outlook limitation of the
      classic (non-VSTO) `CustomTaskPane` docking API when combined with the
      newer nav rail, which classic CTPs predate and may not have been
      updated to account for; the API exposes no direct position/bounds
      control beyond `DockPosition` + `Width`/`Height`. May be specific to
      this ARM64/Prism-emulated dev machine — re-check on the real 32-bit
      target before investing further. Cosmetic only; does not block using
      the pane in its default Right-docked position.
      **Research findings (2026-07-10, web research session):** the icon rail
      is Outlook's "app bar" / Hub Bar feature (feature flag
      `Microsoft.Office.Outlook.Hub.HubBar`), added to classic Outlook in
      Version 2211 (Nov 2022) — some 15 years after the classic CTP docking
      API, which was never updated to account for it. The rail overlapping
      third-party docked panes is a *known* complaint on ordinary x86/x64
      machines (e.g. iManage's Outlook integration; Slipstick Systems
      documents the rail "cuts into the full display of Outlook" and cannot
      be narrowed) — so this is an Office-side layout limitation, **not**
      ARM64/Prism-specific, and should be expected to reproduce on the 32-bit
      Semi-Annual target as well (SAC builds have carried the rail since
      2023). **No add-in-side code fix exists:** the classic
      `_CustomTaskPane` COM interface has no position/bounds members, and the
      only P/Invoke prior art (FindWindow `MsoCommandBar` + `MoveWindow`)
      applies to *floating* CTPs only — for docked panes Outlook's docking
      manager owns the geometry and re-applies it on every window resize, so
      a `SetWindowPos` offset would be immediately undone (and raising the
      CTP's z-order would merely hide the rail's icons instead — no better).
      **Working user-level workaround** (verifiable on this dev machine):
      File → Options → Advanced → "Outlook panes" → uncheck **"Show apps in
      Outlook"**, restart Outlook — moves the navigation back to the bottom
      of the folder pane, removing the rail and the overlap entirely. This is
      the same workaround the iManage community uses. The checkbox is still
      present in recent builds (confirmed in M365 Version 2502) but Microsoft
      has removed it for some users/builds; the registry equivalent —
      `Microsoft.Office.Outlook.Hub.HubBar` = `"false"` (REG_SZ) under
      `HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook`
      or `HKCU\Software\Microsoft\Office\16.0\Common\ExperimentEcs\Overrides`
      — is documented (Slipstick) but flagged as an experimental override
      that may stop working in future builds. **Disposition:** treat as a
      known, unfixable-in-code Office limitation; keep the default Right
      dock; if a pilot user insists on Top/Bottom docking, offer the "Show
      apps in Outlook" workaround (or its registry override where the
      checkbox is absent) rather than attempting a code change.
      Sources: slipstick.com/outlook/outlooks-left-navigation-bar/,
      ablebits.com/office-addins-blog/outlook-customize-move-navigation-pane/,
      aritrasaha.wordpress.com/2009/05/19/ (floating-CTP MoveWindow prior art).
      **RESOLUTION (2026-07-11).** The "no add-in-side fix" disposition was
      half-right: the *pane window* indeed cannot be moved, but the overlap is
      fixable from inside it. Shipped fix, two parts:
      (1) **Content inset** — `NavRailInset.MeasureLeftOverlapPixels`
      (`RBLclass.AddIn/NavRailInset.cs`) geometrically detects any visible
      narrow (≤150px), tall (≥⅓ Explorer height) direct child of the Explorer
      frame covering the host's left edge (the rail is a `NetUINativeHWNDHost`
      'NUIDocumentWindow' strip, ~48 DIP wide, spanning the full client height
      *over* the docks) and `RblClassTaskPaneHost` applies the measured overlap
      as left `Padding` around the `ElementHost` (recomputed, coalesced, on
      every `Resize` + `RefreshOnShow`). Measures 0 — no padding — when docked
      Left/Right, floating, or the rail is hidden, so it is self-adapting.
      (2) **Title indent** — the Office-drawn CTP title strip cannot be padded
      (Office renders it) and clipped "RBLclass" down to "ass"; the title is
      fixed at `CreateCTP` time, so `EnsureTaskPane` prepends 14 spaces when
      the persisted dock position is Top/Bottom (unconditional: the rail HWND
      does not exist yet that early in startup, and a rail-less setup just gets
      a harmless indent). A live redock to Top/Bottom keeps the unindented
      title until the next restart.
      **Hard-freeze regression, root-caused and fixed the same day:** the
      earlier `DockPositionStateChange` → `ApplyDockSize` resize (6f3d723 /
      621ebb3) froze Outlook whenever the pane was live-docked. Two distinct
      failure modes observed: inline `Width`/`Height` set → Office rejects with
      0x80004005 *"property cannot be set while the object's event handler is
      executing"* and, deferred past the guard (`Dispatcher.BeginInvoke` at
      `ApplicationIdle`, even with the event handler detached and value-guarded
      sets), the set *wedges Outlook's docking manager instead*: message pump
      stays alive (`Responding=True`, idle CPU, AppHang only in the inline
      variant) but the whole UI stops taking input until the process is killed.
      Differential evidence: Left/Right/Floating live redocks never froze
      (width already 360 → value guard skipped the COM set); Top/Bottom froze
      every time (height differed → set executed). **Standing rule: never set
      `CustomTaskPane.Width`/`Height` after creation.** `ApplyDockSize` runs
      only from `EnsureTaskPane` (pane not yet visible, handler not yet
      attached — proven safe); the dock-change handler persists the position
      and nothing else. Trade-off: a live redock keeps whatever size Office
      assigns until the next Outlook start. Verified interactively 2026-07-11
      (repeated Top/Bottom/Left/Right/floating redocks, no freeze; content and
      title clear of the rail on a Top-docked cold start).

### Dropped from tracking (user decision, 2026-11-21)

- Banner strip on reply and on classify: duplicate of v2.9.0.0 B1, still
  tracked in ROADMAP.md [#46].
- Intermittent `MailItem.Move` failure: dev Gmail/IMAP account only, points
  at IMAP sync state; not reproducible as a code defect.
- Delete key ignored after pane focus: issue #49 closed; workaround is a
  click in the message list.
- MSI "installer is no longer responding": issue #66 closed with no field
  log; the Outlook lock alone did not reproduce it on the dev machine.

The original item text, for the record:

- [ ] **Strip the external-sender banner on reply and on classify** —
      implemented in the v2.1.0.0 scope; didn't work, so was re-implemented in v2.6.0.0 - pending tests
- [ ] **Intermittent `MailItem.Move` failure (0x80020009) on the dev Gmail/IMAP
      account** — classify into a valid destination fails for *some* items with
      "may not accept items" (per-item, not per-folder; other items to the same
      folder succeed). Surfaced during v2.7.1.0 F2 testing but predates the
      sprint and is unrelated to any v2.7.1.0 change (mark-as-read is now applied
      only after a successful move, so a failed move no longer changes read
      state; the read-state seen on failure was Outlook's own "mark items as
      read when viewed in the Reading Pane"). Likely a Gmail/IMAP move quirk or a
      stale selection EntryID. Investigate on the real 32-bit target / a
      non-Gmail store before pilot; not reproducible as a code defect on the
      current dev account. 2026-11-07: seen again as `0x80040109` ("the
      message has been modified") on real Gmail mail shown in the Reading Pane
      right after a classify, repeating on the same item for 10+ seconds. Not
      reproduced on 47 dummy mails in a Gmail IMAP folder (read and unread,
      1.2 s to 7 s apart), with or without v2.11.0.0 B3. One dummy mail moved
      out of that folder reappeared in it later (IMAP sync), which points at
      IMAP sync state rather than RBLclass code.
- [ ] **Delete key ignored once the pane has had keyboard focus (#49
      regression).** The v2.9.0.0 B2 fix (`Explorer.Activate()` after
      navigating) no longer works on Outlook 16.0.20527 (Current channel). An
      A/B test showed the defect also exists without v2.11.0.0 B3. Measured
      2026-11-07 with an automated harness (dummy `[RBLCLASS-TEST]` mails in
      dedicated folders, UI Automation, real mouse clicks, `SendKeys`; results
      are only reliable with nobody using the PC), reproducible 3/3:
      - Before any pane interaction, Delete deletes the selected mail.
      - Any keyboard focus in the pane kills it, with or without a click and
        with or without "open folder": after that, Delete does nothing in any
        folder. Arrow keys still move the selection, and the Delete command
        itself works (`CommandBars.ExecuteMso("Delete")` deletes), so only the
        accelerator is lost.
      - A real mouse click on a message row restores it. Programmatic focus
        changes do not: `Explorer.Activate()`, `SetFocus` on the `OutlookGrid`
        window (shipped in B3: it fixes "no window has focus" after "open
        folder", but not Delete), clearing WPF focus, collapsing and expanding
        the navigation pane.
      - Calling `IOleInPlaceObject::UIDeactivate` on the pane's ActiveX host
        after "open folder" returned S_OK but changed nothing (reverted).
        WinForms returns S_OK without acting when it does not consider itself
        UI-active, which fits focus being held by the WPF `HwndSource` child
        rather than the WinForms control, so another object may be the one
        swallowing the accelerator.
      Suspected mechanism, not proven: while a hosted control is Office's
      active object, Office offers it every keystroke through
      `IOleInPlaceActiveObject::TranslateAccelerator`, and WinForms forwards
      keys aimed at other windows to the Office control site
      (`Control.ActiveXImpl.TranslateAccelerator`), which can consume them.
      Workaround meanwhile: click in the message list before pressing Delete.
      Next steps: find which object Office routes accelerators to once the
      pane has had focus; evaluate Microsoft's KB 4490421 workaround (an
      intermediary borderless Form between the CTP UserControl and the
      ElementHost, documented for keys not working as expected with WPF in a
      CustomTaskPane), mindful of the "never resize the CTP after creation"
      rule and the nav-rail inset logic; re-verify on the 32-bit target. [#49]
- [ ] **Fix the MSI "installer is no longer responding" failure** (requires:
      v2.11.0.0 A2, plus a real log or a positive repro — blocked on field
      data, so kept out of the sprint section). Root-cause analysis so far
      (2026-11-05): the failure is not in RBLclass code, the package has no
      custom action and no service, only files and HKCU keys. Leading
      hypothesis: Outlook open during a major upgrade, holding
      `RBLclass.AddIn.dll` and `e_sqlite3.dll` while
      `MajorUpgrade Schedule="afterInstallInitialize"` removes the old files,
      with no running-Outlook check and no files-in-use UI in the package.
      Alternatives: EDR/antivirus slowing `msiexec` on an unsigned package;
      Windows Installer service busy or degraded. None is proven. Apply the
      fix the log points to. If it is the Outlook lock, a pre-install check
      probably needs the `WixToolset.Util` extension — a new build
      dependency, so ask before adding it. [#66]
      Discriminating test (2026-11-06, dev machine, 64-bit Outlook, no EDR):
      2.9.1.0 installed, Outlook open and holding the DLLs, 2.10.0.0 MSI run.
      Windows Installer shows its built-in modal "close the following
      applications" dialog after 2 s and waits; the package does get a
      files-in-use UI. Choosing automatic close finished in 3 s (exit 0,
      Outlook restarted). So the lock alone does not reproduce the symptom
      here; this does not clear the target (32-bit Outlook, EDR).

### Phase 9 — completed hardening items

- [x] **Threat-model / security architecture doc** (`docs/security.md`): trust
      boundaries, data handled (folder names/paths + transient send-time
      metadata; **no mail bodies persisted**), data-at-rest location
      (`%LocalAppData%\RBLclass\`, per-user ACL), "no telemetry / no network by
      default", HKCU-only / no-admin posture. **Created (v2.7.0.0)**; the
      network-posture section documents the E1 opt-in update check as the single
      outbound call (off by default, no data sent beyond the GET).
- [x] **Confirm & document log hygiene:** assert no recipient addresses or
      mail-body text reach the Serilog sink. Capture the redaction policy in
      `docs/security.md`. **Documented (v2.7.0.0)** in the "Log hygiene" section
      of `docs/security.md`.
