# RBLclass roadmap

**Closed history** (Phases 0–5, the shipped v2.0/v2.1/v2.2 waves, and the
already-shipped v2.4.0.0 quick-wins) now lives in
[docs/roadmap-history.md](docs/roadmap-history.md). This file holds **where we
are**, the **active v2.4.0.0 sprint**, and the **open forward-looking phases**.

Reference docs from the legacy analysis:
[docs/legacy-overview.md](docs/legacy-overview.md),
[docs/legacy-reimplementation-roadmap.md](docs/legacy-reimplementation-roadmap.md),
[docs/parity-and-regression.md](docs/parity-and-regression.md).

## Where we are (2026-06-20)

The rewritten add-in reached full legacy parity and shipped as **v2.0.0.0**,
followed by two operational feedback waves (**v2.1.0.0**, **v2.2.0.0** — all in
[history](docs/roadmap-history.md)). The **v2.4.0.0** sprint is in progress: the
quick-wins / pane-redesign / results-display items (A–C) have shipped; this file
now carries the **bug-fix and new-capability work** added to that sprint —
classify performance/freeze, classify-after-send, and the attachment-disposition
capability — plus the still-open Stormshield-cosmetic and external-banner items.

Phase 6 (pilot rollout) remains in progress; Phases 7–9 are forward-looking.

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
  increments are re-checked on the real 32-bit target before pilot. Several
  v2.4 items below are explicitly **not verifiable on the dev machine**
  (Stormshield/S-MIME mail, the external banner) and carry a re-verify note.

---

## v2.4.0.0 — Third feedback wave (active sprint, started 2026-06-20)

Already shipped in this sprint (detail in
[history](docs/roadmap-history.md#v2400--shipped-items-third-feedback-wave)):
**A1** clear search after classify · **A2** async index refresh + colour
indicator · **A3** dock-position persistence · **B1** Options-toggle collapse ·
**C1** per-row expandable folder path.

The remaining work, sequenced so the **performance/freeze fix (D2) lands first**
(every later item files mail and benefits from a responsive classify), then the
**after-send** convenience (E, small, reuses triage), then the **large
attachment-disposition capability** (F) last. Every task ends with a build,
`/reload-addin`, and a manual verification pass before the next starts.
Non-trivial `RBLclass.Core` logic ships with xUnit per CLAUDE.md. Decisions
taken with the maintainer on 2026-06-20 are baked into each item.

### D. Classify reliability & performance

- [x] **D1. MAPI_E_NOT_FOUND on multi-item classify (Stormshield) — cosmetic.**
      The pure-move classify of v2.2 did not eliminate the
      `IMessage.GetAttachmentTable: MAPI_E_NOT_FOUND` emitted by Stormshield's
      `Arkoon.SecurityBox…OnBeforeReadAsync`. The error is **cosmetic** —
      classify succeeds, mail lands — but Stormshield raises a confusing
      notification. It manifests only when **more than one email is selected**.
      Working hypothesis: with multiple items, COM message pumping during the
      second `Move()` lets Stormshield's async scan of item[0] interleave before
      item[0] is fully committed in the destination store.
      - Timing markers are already in place (`OutlookMailStore.MoveItemToFolder`
        logs `Classify move BEGIN/END (D1)` with elapsed ms).
      - **This item now folds into D2:** D2's inter-item pump-yield is the exact
        remedy the original D1 candidate fix proposed (release the STA for one
        COM event cycle after each move so Stormshield's scan completes against
        a stable item). Verify on the Stormshield 32-bit target that the
        responsive-classify yield from D2 also clears this notification; if a
        residual remains, add the optional hidden `InterItemDelayMs` fallback.
        Close D1 against the same target pass that validates D2.

- [x] **D2. Classify is slow and freezes the whole Outlook UI — investigate,
      then make classify responsive.** In production, classify sometimes takes
      several seconds and Outlook is **fully frozen** for the duration, despite
      the "optimistic / async" intent in CLAUDE.md.

      **Root cause of the freeze (confirmed by reading the code):**
      `MainPaneViewModel.DoClassify` calls `_classifier.Classify(...)`
      **synchronously on the WPF/STA UI thread**
      ([MainPaneViewModel.cs:522](src/RBLclass.AddIn/ViewModels/MainPaneViewModel.cs#L522)).
      The lone forced `DispatcherPriority.Render` paints the busy state once
      (line 503), then the message pump is **blocked for the entire run** — no
      repaint, no input, Outlook appears hung. Classify is not actually async
      anywhere today.

      **Investigation steps (find the delay source before the fix):**
      1. Add `Debug`/`Information` Serilog timing markers around each phase of a
         classify run — `GetSelectedItems`, `Preflight` (note: conversation
         widening calls `GetConversationSiblings` → `Conversation.GetTable`,
         which can be slow on large mailboxes), each per-destination
         `Copy`/`Move`, `RemoveAttachments`, `ClassificationHistory` writes,
         the safety-copy move. (Move markers already exist from D1.)
      2. Reproduce across the axes that change the work done: 1 vs N items,
         widen-conversation on/off, keep-copy on/off, remove-attachments on/off,
         safety-copy on/off. Identify which phase dominates the multi-second
         cases.
      3. Confirm whether the dominant cost is our COM round-trips, Stormshield's
         synchronous `OnBeforeRead` scan running inline per item, or the
         on-UI-thread SQLite history writes.

      **Fix — responsive classify with progress (decision 2026-06-20: keep moves
      on the STA but never freeze the pump; the user waits with feedback rather
      than the pane returning optimistically).** COM must stay on the STA, so
      this is about **pumping the message loop**, not background threading:
      - Make the classify path `async Task`. Drive the per-item loop so that
        **after each item is moved/stripped/labelled, the STA message pump is
        yielded** (`await Dispatcher.Yield(DispatcherPriority.Background)` /
        an STA `DispatcherFrame`) so Outlook repaints and stays interactive
        between items. This also gives Stormshield's async scan a stable window
        (see D1).
      - Report **per-item progress** in the pane status line
        (`Status_Classify_Progress`, e.g. "Filing 3 / 12"), updated as each item
        completes.
      - Move the **non-COM work off the UI thread**: dispatch
        `ClassificationHistory` writes and undo-plan persistence to the thread
        pool (the COM moves themselves stay on the STA).
      - Keep the existing `IsBusy` guard: the pane stays disabled (no
        double-fire) until the run completes, then re-enables. The
        anti-double-fire contract from v2.1.0.0 B is preserved.
      - **Layering:** `RBLclass.Core` must stay UI-free, so model the
        pump-yield as an **injected callback** the shell supplies — e.g.
        `ClassifierService.ClassifyAsync(request, IProgress<ClassifyProgress>,
        Func<Task> yieldBetweenItems)` — with a no-op yield in tests so the
        orchestration stays portable and unit-testable. (Alternative if it
        proves simpler: keep `ClassifierService` synchronous per-item and drive
        the loop + yield from `MainPaneViewModel` over a new single-item
        classify primitive. Pick the smaller change that keeps Core testable.)
      - **Localization:** progress string in `Strings.resx` / `.fr.resx` /
        `.de.resx`, with `_One`/`_Other` plural if phrased per item.
      - **xUnit:** the progress callback fires once per processed item; the
        yield callback is invoked between items; item order and the undo plan
        are unchanged from the synchronous path; encrypted-skip behaviour
        unchanged.
      - **Verification:** measured non-freeze on a multi-item classify on the
        dev machine; re-verify the Stormshield interaction on the 32-bit target
        (shared with D1).

### E. Classify after send

- [x] **E1. Re-offer classification of a just-sent mail via Move-to-Inbox.**
      Some users want to classify a mail right after sending it (a capability
      dropped earlier). **Decision (2026-06-20): the lightweight reuse of the
      existing sent-item triage**, not a new modal. When the triage action
      resolves to **"Move to Inbox"** (either the fixed dropdown value or chosen
      in the "Let me choose" modal), the moved mail is additionally **handed to
      the main pane as the active classify target**, and the pane is
      revealed/focused — so the user keyword-searches and files it immediately
      with the normal pane flow.
      - **Why not Outlook's explorer selection:** the Outlook OM has no reliable
        API to programmatically set `Explorer.Selection` to a specific item, so
        we do **not** try to select it in the message list. Instead the moved
        `MailItemRef` (returned by the existing `MoveItemToFolder` to the Inbox)
        is **pushed into `MainPaneViewModel` as a pinned target** that overrides
        `GetSelectedItems` for the next classify. The pin clears after that
        classify completes, or when the user changes the real explorer
        selection. A small pane indicator shows it is classifying the just-sent
        mail.
      - New setting `ClassifyAfterMoveToInbox` (bool, default **on**) in the
        Settings "Sent items" group, so the behaviour can be turned off.
      - **Layering:** the pin is pane state (`RBLclass.AddIn`); the existing
        triage → move path (`RblClassAddIn` + `IMailStore.GetInboxFolder` /
        `MoveItemToFolder`) is reused unchanged. No `RBLclass.Core` logic beyond
        the new setting key/default.
      - **Localization:** pane indicator + the new setting label/description in
        EN/FR/DE.
      - **xUnit:** setting default/round-trip; pin-overrides-selection and
        pin-clears-after-classify logic if any of it lands in a testable Core/VM
        seam (otherwise shell wiring, verified live — fully reproducible on the
        dev machine: send a mail → triage → classify).

### F. Attachment disposition & filesystem favourites (largest slice)

When mail is classified with **remove attachments** on, instead of silently
deleting them, list every attachment and let the user **delete** it or **save**
it to a fast-to-pick filesystem location, then label the mail with what happened
to each. Sequenced as three sub-items (F1 the searchable favourites index, F2
the disposition modal + save/delete, F3 the label), built in that order.

- [x] **F1. Keyword-searchable "favourite folders" filesystem index.** Saving
      an attachment to the right Windows directory is the slow step;
      pre-indexing user-chosen favourites makes target selection as fast as the
      Outlook folder search. **Decision (2026-06-20): favourites are Windows
      filesystem directories (local/network); the user may also browse to an
      arbitrary directory as a fallback.**
      - **Settings:** a new "Favourite folders" editor in the Settings window —
        add directories via a folder-browse dialog, remove them; persisted as a
        new setting (`AttachmentFavoriteFolders`, path list). Recursive
        sub-directory indexing with a sane depth/count cap (32-bit memory
        budget — index **paths only**, never file contents).
      - **Index + search:** mirror the Outlook folder-search design — persist
        the favourite tree in SQLite (**schema v3**, e.g. a `FavoriteFolders`
        table keyed by path) and search it with the **same matcher**
        (`FolderSearchService` / `FolderMatchMode`, AND across keywords,
        case/accent-insensitive). Re-index when the favourites setting changes
        and on startup.
      - **Layering:** the matcher and the searchable model live in
        `RBLclass.Core` (testable, no I/O); the actual directory enumeration is
        behind an injected interface (e.g. `IDirectoryScanner` /
        `IFavoriteFolderSource`) implemented in the shell, since `RBLclass.Core`
        does no I/O except via injected interfaces and the directory walk is not
        Outlook-specific (so it does not belong in the Outlook adapter).
      - **Localization:** the Settings editor labels in EN/FR/DE.
      - **xUnit:** favourite-folder keyword matching (reuse the
        `FolderSearchService` test patterns), path-segment splitting, depth/count
        cap, empty/duplicate path handling.

- [x] **F2. Per-attachment disposition modal.** **Decision (2026-06-20): a new
      setting picks the behaviour of the remove-attachments checkbox** —
      `AttachmentRemovalMode` = **Modal** (default) or **DeleteSilently** (the
      current no-prompt strip,
      [OutlookMailStore.RemoveAttachments](src/RBLclass.Outlook.Adapter/OutlookMailStore.cs#L742)).
      When Modal and remove-attachments is on, classify shows the modal instead
      of silently stripping.
      - **Flow:** gather attachments across all selected mails (grouped per
        mail). **Encrypted mail (`IPM.Note.SMIME`) is excluded entirely** —
        its attachments always travel with the mail even when the box is
        checked (existing rule, [[never-strip-encrypted-attachments]]) — and the
        excluded items are reported in the modal. Each attachment row shows
        filename, size, and source-mail subject, with a per-row disposition:
        **Delete** or **Save to…** (a folder picker over the F1 favourites
        index with keyword search + a browse-to-arbitrary fallback). Provide a
        sensible default + "apply to all" convenience.
      - **On confirm:** save each "Save" attachment to its chosen directory
        (resolve filename collisions by appending ` (1)`, ` (2)`…), then strip
        the chosen attachments from the mail and apply the F3 label. Saving and
        stripping run on the STA (COM attachment access); the disposition step
        slots into the classify pipeline **before the move**, operating once on
        the item, and the strip/label follow the existing copy-vs-original
        semantics ([[classify-attachments-on-copy]]: act on the filed copy /
        moved item, leave the original's attachments intact when keep-a-copy is
        on).
      - **Adapter (`IMailStore`):** add
        `IReadOnlyList<AttachmentInfo> GetAttachments(MailItemRef)` and
        `bool SaveAttachmentToFile(MailItemRef, attachmentId, destPath)`;
        `AttachmentInfo` is an Outlook-free Core type (filename, size, id).
        Reuse `RemoveAttachments` for deletion (already skips S/MIME).
      - **Sequencing with D2:** the modal is interactive, so it runs in the
        preflight (before the responsive move loop); the per-item progress of D2
        covers only the move phase.
      - **Localization:** modal chrome, disposition labels, the excluded-encrypted
        notice, and the `AttachmentRemovalMode` setting in EN/FR/DE.
      - **xUnit:** filename-collision resolver (pure Core helper); disposition
        orchestration over a fake `IMailStore` (save vs delete counts, encrypted
        items excluded and reported, label target follows keep-copy); setting
        default/round-trip.
      - **Open point (target re-verify):** Windows **Controlled Folder Access**
        (ransomware protection) blocks Outlook from writing attachments into
        protected directories (Documents, Desktop, OneDrive known folders),
        surfacing as a Windows "unauthorised changes blocked" toast and a
        `SaveAsFile` `FileNotFoundException`. The save-then-strip path now
        **fails safe** — a blocked save leaves every attachment on the mail
        (nothing lost) and the run reports a failure — but saving into a
        protected folder needs CFA to allow `outlook.exe`. Confirm the policy /
        allow-listing on the 32-bit target before pilot.

- [x] **F3. "Former attachments" label on the filed mail.** After disposition,
      record on each affected mail what its attachments were and where each
      went. **Decision (2026-06-20): the label location is a user setting,
      `AttachmentLabelLocation` = `InfoBar` or `Body`.**
      - **Body** (reliable path): append/insert an HTML note block into
        `HTMLBody` listing each former attachment and its disposition ("Saved to
        `<path>`" or "Deleted on `<date>`"), reusing the body-rewrite pattern
        from `StripBannerOnLiveMail` / `ExternalBannerStripper` (never on
        encrypted mail — moot here, they are excluded from F2). Applied to the
        filed copy / moved item, consistent with the strip target.
      - **InfoBar** (`OlkInfoBar`): the maintainer wants the native Outlook info
        bar. **Feasibility spike required first:** `OlkInfoBar` is an Outlook
        **form-region control**, not a property writable on an arbitrary stored
        item, so a persistent per-item info-bar label may need a custom form
        region (or may not be achievable for stored items at all). Investigate;
        if it cannot carry a durable per-item label, **fall back to Body** and
        record the InfoBar option as deferred. Do not block F2/F3 on it — ship
        Body first.
      - **Label composition** is a pure `RBLclass.Core` formatter
        (`AttachmentLabelFormatter`): attachment names + per-item disposition +
        target path/date, with count-dependent `_One`/`_Other` plurals.
      - **Localization:** label template strings + plural pair in EN/FR/DE.
      - **xUnit:** label text for saved-only / deleted-only / mixed dispositions,
        single vs multiple attachments (plural), path and date rendering.

### Carried over from v2.2.0.0 / v2.1.0.0

- [ ] **Strip the external-sender banner on reply and on classify** —
      implemented in the v2.1.0.0 scope; **not verified live** (no
      external-banner mail on the dev machine). Re-verify on a workstation that
      receives the company banner before closing.

---

## v2.4.1.0 — Attachment & indexing bug fixes

- [x] **A1.** Fix the re-index button to fully offload PST enumeration to a background thread so the Outlook UI stays responsive during manual reindex, with the colour indicator cycling to yellow as designed in A2. [#6]
- [ ] **A2.** Show the attachment-disposition modal during auto-classify when "remove attachments" is on, matching the manual-classify flow. [#7]
- [ ] **A3.** Replace the attachment-disposition modal with a discrete status-area notice when classifying encrypted mail (which always retains its attachments). [#8]

---

## v2.5.0.0 — Attachment refinements & UX polish

### A. Status & notifications

- [ ] **A1.** Replace the single status-area notification with a rolling in-session log capped at ~5 entries, so short-lived messages remain visible. [#9]

### B. Attachment enhancements

- [ ] **B1.** Add a "Keep" disposition option per attachment row in the removal modal, leaving chosen attachments on the filed copy untouched. [#5]
- [ ] **B2.** Auto-exclude inline/embedded images from the attachment-removal flow so only true detached files are presented for disposition. [#4]

### C. Packaging

- [ ] **C1.** Update the MSI to detect and remove files from a previously installed version before laying down the new one, eliminating the manual uninstall step. [#10]

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
