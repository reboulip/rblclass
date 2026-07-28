# RBLclass roadmap

**Closed history** — Phases 0–5 and every shipped feedback wave
(**v2.0/v2.1/v2.2/v2.4.0.0/v2.4.1.0/v2.5.0.0**) — lives in
[docs/roadmap-history.md](docs/roadmap-history.md). This file holds **where we
are**, the few **open carried-over items**, and the **forward-looking phases**.

Reference docs from the legacy analysis:
[docs/legacy-overview.md](docs/legacy-overview.md),
[docs/legacy-reimplementation-roadmap.md](docs/legacy-reimplementation-roadmap.md),
[docs/parity-and-regression.md](docs/parity-and-regression.md).

## Where we are (2026-07-11)

The rewritten add-in reached full legacy parity and shipped as **v2.0.0.0**,
followed by feedback/refinement waves **v2.1.0.0** through **v2.8.0.0**
(pane actions, sent-items handling, classify reliability, logging audit) —
all complete and recorded in [history](docs/roadmap-history.md). **v2.9.0.0**
(UI polish, classify & banner fixes, attachments) shipped 2026-07-11 —
[release](https://github.com/reboulip/rblclass/releases/tag/v2.9.0.0).
**v2.9.1.0** (nav-rail freeze hotfix) shipped 2026-07-11 —
[release](https://github.com/reboulip/rblclass/releases/tag/v2.9.1.0).

Open work: the two v2.6.0.0 items still pending live verification (banner
strip B3, meeting classify C1), and v2.9.0.0 B1 (banner strip on
reply/forward) likewise shipped pending live verification on the real
target — **confirmed still broken on the real target on 2026-07-28**
(issue #46: detection works, stripping doesn't); v2.10.0.0 D1 adds the
diagnostic capability needed to root-cause it. **v2.10.0.0** (folder
context-menu completeness, pane consolidation, selection-lag fix, banner
diagnostics) is planned from the 2026-07-28 `/gh-issues` triage. Phases 8–9
are forward-looking.

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
      current dev account.
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

---

## v2.6.0.0 — UI refresh & classify depth

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

---

## v2.9.0.0 — feedback wave: UI polish, classify & banner fixes, attachments

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

---

## v2.10.0.0 — feedback wave: folder management, pane consolidation, performance, diagnostics

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

- [ ] **B1.** Move the "Settings" and "Update folders" actions from the
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

- [ ] **C1.** Fix the long delay in the pane's "N selected" indicator on
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
      FR/DE. [#46]

---

## Phase 6 — Pilot rollout (in progress)

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

## Phase 7 — General availability

- [x] Communication plan.
- [x] Self-service install page (link to the per-user MSI on the internal HTTPS
      share).
- [x] Support process.
- [x] Monitoring of error logs across the fleet.

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
- [x] **Threat-model / security architecture doc** (`docs/security.md`): trust
      boundaries, data handled (folder names/paths + transient send-time
      metadata; **no mail bodies persisted**), data-at-rest location
      (`%LocalAppData%\RBLclass\`, per-user ACL), "no telemetry / no network by
      default", HKCU-only / no-admin posture. **Created (v2.7.0.0)**; the
      network-posture section documents the E1 opt-in update check as the single
      outbound call (off by default, no data sent beyond the GET).
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
- [x] **Confirm & document log hygiene:** assert no recipient addresses or
      mail-body text reach the Serilog sink. Capture the redaction policy in
      `docs/security.md`. **Documented (v2.7.0.0)** in the "Log hygiene" section
      of `docs/security.md`.

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
