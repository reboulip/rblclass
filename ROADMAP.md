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

## v2.7.0.0 — feedback wave: shortcuts, sent-items, docs

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

---

## v2.7.1.0 — code-review remediation sprint

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

---

## v2.8.0.0 — feedback wave: pane actions, sent-items, classify reliability

### A — Pane actions
- [x] **A1.** Add a "save as" button to the pane, next to the Auto-Class button, letting the user save the selected email(s) as files, reusing the attachment-disposition modal's directory picker for the target folder. [#40]
- [x] **A2.** Add a manual attachment-removal button: opens the attachment-disposition modal for the selected email(s) without a classify move, so keep/delete/save-to can be applied standalone. [#38]
- [x] **A3.** In the standalone attachment-management modal (A2's 📎 button), add a "remove saved attachments from the email" checkbox that controls whether a "Save to" attachment is stripped from the email after saving (ON = save then strip, OFF = save but keep it on the email). The checkbox default comes from a new setting in Settings → Attachments. S/MIME items are never stripped. The classify-time modal is unaffected.

### B — Sent-items handling
- [ ] **B1.** Add a "delete" button to the post-sent classify notification, to allow one-click deletion of the sent-and-moved-to-Inbox copy. [#39]

### C — Classify reliability
- [ ] **C1.** Fix auto-class + attachment-modal conflict: auto-class must check for a valid target folder in the conversation history *before* opening the attachment-disposition modal, so the modal never opens for an email it cannot process; when multiple emails are selected, only show attachments for the ones that will actually be processed; surface a clearer (non-modal) message when auto-class cannot process an element. [#36]
- [ ] **C2.** Investigate and fix the classify silent-failure/freeze reported on v2.6.0.0: ~30s "Classement en cours" then no move, reproducible after reinstall/reboot; log attached on the issue. Reporter notes meeting-invite classify is unaffected — only regular mail. [#35]

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
