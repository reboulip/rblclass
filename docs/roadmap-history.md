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
