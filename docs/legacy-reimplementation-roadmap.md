# Legacy reimplementation roadmap

A **sequential, feature-driven** plan to rebuild the legacy RBLclass
feature set (see [legacy-overview.md](legacy-overview.md)) on the target
architecture (COM Shared Add-in + `RBLclass.Core` / `.Outlook.Adapter` /
`.AddIn`, SQLite, WPF task panes — see [CLAUDE.md](../CLAUDE.md)).

This is a **proposal to be challenged**, not an instruction to start
coding. It complements the high-level phase plan in
[ROADMAP.md](../ROADMAP.md) by going feature-by-feature and by calling
out, explicitly, every place where a faithful port would fight our
chosen stack — those are tagged **⚠ DEVIATION** with a recommendation.

## 0. Ground rules carried over from the validated POC

The POC was validated end-to-end on the real 32-bit target. The
reimplementation must not regress any of these load-bearing facts:

- **Loading mechanism = HKCU COM registration**, written by a per-user
  installer. No VSTO, no ClickOnce. The same AnyCPU binary loads on
  64-bit (dev) and 32-bit (target) Outlook.
- **No code signature is required for the add-in to load.** The POC
  installed and ran via an *unsigned* PowerShell installer with the
  workstation's existing execution policy, and Stormshield/EDR did not
  object. → **Authenticode signing stays in the plan as defence-in-depth
  and as a Phase-1 *option*, but it is not a functional gate.** Treat
  any signing task as "nice to have / do if the internal PKI is ready",
  never as a blocker for a working build. This is the one place the
  current [ROADMAP.md](../ROADMAP.md) is too strong (it lists ClickOnce
  and signing as Phase-0 gates); reconcile it toward "PowerShell
  installer for POC → optional-signed WiX MSI for GA".
- **SQLite native payload must ship both `win-x86` and `win-x64`**
  `e_sqlite3.dll`, verified in build output. Already proven in the POC.
- **Canonical GAC PIAs only** for Extensibility/Office/Outlook interop;
  never hand-rolled `[ComImport]`. Already proven (and scarred) in the
  POC.

Everything below assumes the Phase 1 solution skeleton from
[ROADMAP.md](../ROADMAP.md) (the three projects, `ComRef<T>`, Serilog,
SQLite migrations, WiX skeleton) already exists. The steps are ordered
so that each builds on the previous and each ends at a **demoable
increment**.

---

## A key scoping insight before sequencing

The legacy tool **never indexed mail bodies** — its only index is the
**folder tree**, and its keyword search runs over **folder path
strings**. Therefore:

> **Legacy feature parity does NOT require the FTS5 mail index** or the
> "<100ms over 100k mails" target from ROADMAP Phase 2. Those serve
> *future* smart features (auto-classification suggestions), not the
> legacy behaviour.

**⚠ DEVIATION (scope/sequencing):** I recommend we **defer the mail
FTS5 index** out of the parity track entirely and build only a
**folder index** first. The folder index is small (thousands of nodes,
not 100k mails), so its search does not need FTS5 — see Step 2. This
de-risks and shortens the path to a usable tool. The mail index returns
later, only when we start the smart-suggestion features.

---

## Testing & validation — built into every step

Two complementary means, present from Step 1 onward so nothing is built
without a way to check it:

1. **Automated unit tests (`RBLclass.Core.Tests`, xUnit +
   FluentAssertions + NSubstitute).** Because all decision logic lives in
   `RBLclass.Core` behind interfaces, it is testable without Outlook:
   - **Folder search (Step 2)** — the highest-value surface. Port a
     corpus of folder paths and assert match / collapse / expand /
     ordering / `MaxResults` cap against the documented legacy semantics,
     including the substring cases (`"PJ"` → `ProjetJuridique`) and
     accent/space-insensitivity.
   - **Classifier (Steps 4, 6)** — drive `IClassifier` with a faked
     `IMailStore` (NSubstitute); assert copy-per-destination,
     delete-vs-keep, conversation widening, and task-completion decisions
     without touching real mail.
   - **Send guards (Step 7)** — pure functions (body text + recipients +
     attachment count → decision); table-driven tests over the keyword
     list and internal/external rules.
   - **Settings (Step 9)** — round-trip and defaults over an in-memory
     SQLite DB.
   These run in **CI on every push**. The Adapter and AddIn are excluded
   from CI (they need a live Outlook).

2. **An installable add-in on this dev workstation, every step.** The
   POC PowerShell installer is promoted to a reusable
   `Install-RBLclass.ps1` / `Uninstall-RBLclass.ps1` (HKCU registration,
   no admin, no signature — exactly the validated Phase-0 mechanism), and
   is a **first-class build output**. Every step above ends at a
   **Demo** line describing what to load into the dev Outlook and click
   to validate the increment by hand before moving on. The 64-bit dev
   Outlook is the fast loop; parity-critical steps also get re-checked on
   the 32-bit target.

> Note on the Adapter layer: it has no automated coverage by design (live
> Outlook dependency). The installable-add-in loop is its test harness —
> which is why the installer is treated as core infrastructure, not
> packaging that can wait until the end.

---

## Step 1 — Folder-tree model & indexing (Core + Adapter)

**Goal:** reproduce `indexFolders` / `FolderInPST` on the new
architecture, decoupled from Outlook.

- **Core** (`RBLclass.Core`):
  - `FolderNode` domain type: `StoreId`, `EntryId`, `Name`, `FullPath`
    (path segments), `ParentId`, `IsLeaf`. Replaces `FolderInPST`.
    Keep `StoreId`+`EntryId` as the stable key (the legacy model kept
    only name strings — a known fragility we fix here, consistent with
    the CLAUDE.md "cache by (StoreID, EntryID), tolerate misses" rule).
  - `IFolderTree` interface: enumerate stores, walk folders, expose the
    cached tree.
  - `IMailStore` interface (the Outlook seam): `GetStores()`,
    `GetFolders(storeId)`, `NavigateTo(folderId)`, `CreateSubfolder(…)`,
    etc. Pure interface, no Outlook types.
- **Adapter** (`RBLclass.Outlook.Adapter`):
  - Implement `IMailStore`/`IFolderTree` over the Outlook OM using
    `ComRef<T>`, index-based `for` loops, explicit release (per CLAUDE.md
    COM-lifetime rules). This is the recursive `archiveExplo`/
    `inDepthExplo` walk, but releasing every intermediate folder.
  - **Exclusions become configuration, not hard-coded French strings.**
    ⚠ DEVIATION (correctness): replace the `"Dossiers publics"` /
    `"G-FRA"` / `"Éléments supprimés"` substring tests with (a) the
    `DefaultItemType`/`FolderType` for Deleted Items where available,
    and (b) a configurable exclusion list (store names / well-known
    folder kinds). Document any DASL used.
- **Persistence — the index rests in SQLite between sessions.** Store
  the folder tree in a `Folders` table
  (`StoreId, EntryId, ParentEntryId, Name, FullPath, IsLeaf`) — the
  relational shape the legacy `tests.bas` Access prototype already
  sketched. Versioned migration `v1`.
- **Index lifecycle (this is the answer to "do we re-index every
  startup?" — no):**
  1. **First run only:** full background walk of the open PSTs → write
     the tree to SQLite.
  2. **Every later start:** **load the tree from SQLite, do NOT re-walk
     Outlook.** Loading a few thousand rows is milliseconds; the legacy
     tool only re-walked on every start because it had no persistent
     store.
  3. **Incremental freshness:** subscribe to Outlook **folder** events
     (`Folders.FolderAdd` / `FolderRemove` / `FolderChange`) plus a
     cheap background reconcile to catch changes made while the add-in
     wasn't running; sub-folder creation re-indexes only that subtree.
     No full rebuild.
  This persistence is exactly the property you wanted to keep — it is
  **orthogonal to deviation #2**, which is only about the *matching
  algorithm* (substring vs FTS5), never about where the index lives.
- **Threading:** the first-run walk and the background reconcile run on
  the thread pool after `OnStartupComplete`; marshal UI updates back via
  the captured `SynchronizationContext`. Replaces the legacy modal
  `pleaseWait` form.

**Demo:** first start indexes all open PSTs to SQLite (logs counts); a
second start loads from SQLite with no Outlook walk; creating a folder in
Outlook shows up in search without a restart.

---

## Step 2 — Keyword folder search (Core)

**Goal:** reproduce `kwSearch` semantics exactly.

- **Core service** `IFolderSearch.Search(string query, bool allResults)`:
  - **AND across whitespace-split keywords**, **substring** match,
    **case-insensitive**, **accent- and space-insensitive** (the legacy
    strips spaces and uses `vbTextCompare`; we additionally normalise
    diacritics so `"reglement"` matches `"Règlement"`).
  - Reproduce the prune/collapse/expand logic: match a folder when all
    keywords are on its ancestor path; collapse a matched non-leaf to
    `path|[...]` unless **All results**; cap at `MaxResults`.
  - Alphabetical ordering of results.
- **Implementation note / ⚠ DEVIATION (matching algorithm only — NOT
  storage):** do **not** use FTS5 to *match*. FTS5 tokenises and would
  break the substring/partial semantics that make `"PJ"` match
  `ProjetJuridique`. Instead, match over the SQLite-backed folder set —
  either by hydrating the cached tree into memory once at startup
  (fastest, matches legacy behaviour 1:1) or via SQL `LIKE '%kw%'`
  AND-chained against the `Folders` table. **The folder index still
  lives in SQLite either way** (Step 1); this deviation changes only how
  we compare keywords, not where the data rests. Reserve FTS5 strictly
  for *mail* search if/when we add it (Phase 5) — keeping us consistent
  with CLAUDE.md ("FTS5 for mail body/subject") rather than
  contradicting it.
- **Tests:** this is the highest-value unit-test surface — port a corpus
  of folder paths and assert match/collapse/expand/ordering/cap against
  the documented legacy semantics. (Core is the only tested layer.)

**Demo:** unit tests green; a console/debug harness returns the same
folder hits the VBA tool would.

---

## Step 3 — Add-in shell, ribbon & first interactive action: Open a folder

**Goal:** first end-to-end user-visible feature = legacy §5a.

- **AddIn** (`RBLclass.AddIn`): `IDTExtensibility2` +
  `IRibbonExtensibility` shell (the POC's shape, promoted to the real
  solution). `[ClassInterface(AutoDispatch)]`, GAC PIAs. Ribbon XML with
  an **"Open folder"** button.
- **WPF task pane** via `ICTPFactoryConsumer`/`ICustomTaskPane`,
  MVVM with CommunityToolkit.Mvvm:
  - Search text box (type-ahead) bound to `IFolderSearch`.
  - Results list; **Enter** = pick first + go, double-click = go,
    a modifier toggles **All results** (legacy used Ctrl — keep or
    rebind, user's call).
  - "New window vs same window" option.
- **Adapter:** `NavigateTo(folderId)` → `Explorer.CurrentFolder` or
  `MAPIFolder.Display`. Re-resolve by **(StoreId, EntryId)** rather than
  re-walking a pipe-path (legacy `setTargetFolder`), tolerating misses.

**⚠ DEVIATION (UX surface):** legacy used floating MSForms modal
dialogs; we use a docked **Custom Task Pane**. Behaviour is preserved,
presentation is modernised. If the user prefers a popup feel for the
"open"/"classify" actions, a modeless WPF `Window` is an option — flag
for the user to choose.

**Demo:** type keywords in the task pane, jump to the folder. First real
parity feature shipped.

---

## Step 4 — Classify selection (Core + Adapter + AddIn)

**Goal:** the flagship action, legacy §5b — the bulk of the value.

- **Core** `IClassifier.Classify(ClassifyRequest)` where the request
  carries: selected item ids, destination folder ids (1..N), and the
  toggles (include conversation, keep copy, remove attachments,
  mark tasks complete). All decision logic lives here and is unit-tested
  with a faked `IMailStore`.
- **Adapter** implements the Outlook mechanics:
  - **Copy-per-destination then Move** (preserve legacy semantics: N
    destinations → N copies). Honour the CLAUDE.md rule that
    `MailItem.Move` invalidates the original reference — capture the
    returned moved item, release the source.
  - Optional **delete originals** when *keep a copy* is off.
  - All over `ComRef<T>`, `for`-index loops, no property chaining.
- **AddIn:** ribbon **"Classify"** button + task pane reusing the same
  search UI as Step 3, but with a **multi-select** result list and the
  per-action toggles.
- **Optimistic UI:** show success immediately, perform the `Move`
  asynchronously (CLAUDE.md perceived-latency target <300ms). On
  failure, surface via Serilog + a non-fatal notice.

**Sub-features folded in here:**
- **Sub-folder creation** (§5c): `IMailStore.CreateSubfolder` + targeted
  re-index of just that store's subtree (not a full rebuild).
- **Attachment removal** (§5e): `IClassifier` honours the toggle;
  also expose a standalone ribbon **"Remove attachments"** command over
  the current selection. Confirmation dialog preserved.

**Deferred to Step 6 (kept out of the first classify cut to ship
sooner):** conversation widening and the task-completion guard.

**Demo:** select mails, classify to one or several folders, originals
deleted or kept per toggle, attachments optionally stripped.

---

## Step 5 — DROPPED: Quick Open popup & "last filed folder"

**Goal (legacy §5d):** after a classify, jump to the last filed folder
via a small popup affordance.

**Status: dropped (2026-06-07).** A pane-local infobar version was built,
tried live, and didn't land well. The user does not want this feature
ported — it will be reconsidered, if at all, only as part of a deeper UI
pass later (see Step 10), as something "more meaningful". Do not
re-attempt this as specified here.

---

## Step 6 — Conversation widening + task-completion guard (Core + Adapter)

**Goal:** legacy §5b steps 2–3, now de-risked.

- **Conversation widening:** add to the classify set every item sharing
  a selected item's conversation across **Inbox + Sent Items**.
  ⚠ DEVIATION (performance/correctness): the legacy code **iterates
  every item** in both folders (O(n) per classify) and compares
  `ConversationID`. Replace with `Items.Restrict`/`Find` using a DASL
  filter on the conversation property, or Outlook's
  `MailItem.GetConversation`/`ConversationID` lookup. Document the DASL.
  Dedupe by `(StoreId, EntryId)`.
  **Implementation note (2026-06-07):** built on `MailItem.GetConversation()`
  → `Conversation.GetTable()`, not DASL/`Restrict` — Outlook indexes the
  conversation table itself, so this is both more robust (no guessing
  binary-MAPI-property DASL syntax) and faster (a small Outlook-side
  result set vs. a full-folder scan). Rows are filtered to the default
  Inbox/Sent EntryIDs and deduped by `(StoreId, EntryId)`.
- **Task-completion guard:** detect task-flagged, not-completed items and
  offer to mark them complete in the destination.
  ⚠ DEVIATION (correctness): drop the `TaskCompletedDate = "01/01/4501"`
  string sentinel; use the proper OM (`TaskCompletedDate` /
  `IsMarkedAsTask` / the `olTaskStatus` semantics) and a real null/not-
  set check, not a locale-formatted magic date.

**Demo:** classify one mail of a thread → its Inbox/Sent siblings come
along; flagged tasks prompt to complete.

---

## Step 7 — Send-time guards (AddIn + Core)

**Goal:** legacy §6a–§6b.

- **AddIn** subscribes to `Application.ItemSend` in `OnStartupComplete`,
  unsubscribes in `OnBeginShutdown`, keeps the source object in a field
  (CLAUDE.md event rules). Top-level try/catch — never let it throw into
  Outlook.
- **Core** holds the *pure* decision logic, unit-tested:
  - **Forgotten-attachment check:** given the latest-reply body text +
    attachment count + a configurable keyword list → bool.
    ⚠ DEVIATION (robustness): the legacy isolates the latest reply via
    the theme-specific HTML separator `solid #B5C4DF 1.0pt`, which is
    brittle. Prefer Outlook's reply/markers or `PropertyAccessor` on the
    body where possible; if we must heuristically split, make the marker
    and the keyword list **configurable** and localised, not hard-coded.
  - **External-recipient check:** ⚠ DEVIATION (correctness): replace the
    `CORPMAIL` substring test with a robust internal/external decision —
    inspect `Recipient.AddressEntry.Type`/`AddressEntryUserType`
    (Exchange user vs SMTP) and/or a **configurable internal-domain
    allowlist**. PST-only environments with no Exchange must still work
    (default everything to "external" → warn, or let config disable the
    check).
- The pile of French `AppActivate` focus hacks is **dropped** — a modal
  WPF dialog owns focus correctly.

**Demo:** sending with "joint" in the body and no attachment warns;
sending to an outside address warns; both toggleable.

---

## Step 8 — Sent-item triage (AddIn + Core)

**Goal:** legacy §6c.

- **AddIn** subscribes to the **Sent Items** `Items.ItemAdd` event
  (keep the `Items` reference in a field). Reproduce the legacy
  **re-entrancy guard** cleanly: instead of `Set … = Nothing` + re-init,
  use a `bool _suppress` flag around the processing block.
- Prompt (WPF) offering **Class / Delete / Move-to-Inbox / Leave**, with
  a **whole-conversation** option, reusing `IClassifier` and the
  conversation-widening from Step 6.

**Implementation note (2026-06-07):** the prompt is a small modal
`Window` with hand-rolled MVVM (`SentItemTriageViewModel` /
`SentItemTriageWindow`) — the first modal dialog in the add-in; every
prior prompt was a `MessageBox`, which can't express four custom actions
plus a toggle. "Class…" opens a second small modal, a dedicated
folder-picker (`FolderPickerViewModel` / `FolderPickerWindow`) over
`IFolderSearch`, deliberately **not** the existing Classify pane —
that pane is built around the live Outlook selection
(`ClassifyViewModel`'s `_getSelection`), and driving it with a fixed,
externally-supplied item list would mean teaching a shipped, working
flow a second selection model for one caller. The orchestration itself
(widen → preflight → optional task-completion confirm → classify/delete/
move) lives directly in the AddIn shell, mirroring `ClassifyViewModel
.DoClassify` — there is no new Core decision logic to test here, only
two small `IMailStore` extraction methods (`ResolveMailItem`,
`GetInboxFolder`) that, like `InspectForSend`, turn a live COM item into
an Outlook-free DTO. "Move to Inbox" is `Classify` with `keepCopy:
false` to a resolved `olFolderInbox` destination — a real move, not a
new code path. The legacy `Set colSentItems = Nothing` / re-init dance
is replaced by a plain `_suppressSentItemTriage` flag set around the
dispatch (a classify/move makes a copy transit Sent Items, which would
otherwise re-trigger the prompt on the transient copy).

**Demo:** send a mail → triage prompt → classify/delete/leave works,
no event re-entrancy.

---

## Step 9 — Settings (Core + AddIn)

**Goal:** legacy §7, modernised.

- **⚠ DEVIATION (storage):** replace `%APPDATA%\…\options.txt`
  positional text with the **SQLite `Settings` table** mandated by
  CLAUDE.md (key/value, typed accessors, defaults, migration). This kills
  the legacy "missing line throws → recreate" brittleness.
- Port the nine options (UFsent, SendExt, AllResults, WholeConversation,
  deleteProcessedElts, NewWindow, RemoveAttachments, QuickOpen,
  MaxResults) as a typed `Settings` object in Core, surfaced through a
  WPF settings pane. Sensible first-run defaults (same as legacy).
- No "force the form on top until saved" hack — defaults are valid out of
  the box, so the app is usable before the user ever opens settings.

**Demo:** toggle options, persist across restarts, each guard/action
respects its setting.

**Implementation note (2026-06-07):** shipped as a modal "Options"-style
dialog (`SettingsWindow`/`SettingsViewModel`, the third modal `Window` in
the project after Step 8's triage/picker dialogs and a clean fit for a
configure-once-and-walk-away flow), opened from a new ribbon button, with
**live-apply, save-on-change** persistence — every toggle/edit writes
through `ISettingsStore` immediately, exactly like the existing
Classify/folder-search panes, so "Close" is the only button and there is
no separate save step to forget.

QuickOpen was already dropped ([[step5-quick-open-dropped]]), so eight
legacy options were ported, plus three settings introduced earlier in the
rewrite that had keys but no UI yet (`FolderMatchMode`, `InternalDomains`,
`ForgottenAttachmentKeywords`) and `MaxResults`, which existed only as the
hard-coded `FolderSearchOptions.DefaultMaxResults` until now — eleven
options in total, grouped into four sections (Folder search / Classify /
Sending guards / Sent items).

The "typed `Settings` object in Core" is `RBLclass.Core.Settings`: a POCO
snapshot with a static `Load(ISettingsStore)` (defaults + parsing +
clamping - invalid `MaxResults`/`FolderMatchMode` values fall back rather
than propagate) and an instance `Save(ISettingsStore)` that round-trips
every key as one unit. The dialog is the only thing that loads/saves the
whole snapshot; the existing scattered `GetBool(SettingsKeys.X, default)`
call sites in `ClassifyViewModel`/`FolderSearchViewModel`/the shell were
left untouched (per [[prefer-isolated-new-ui-over-retrofit]] - don't
retrofit working code) **except** for the three `FolderSearchOptions`
construction sites (`ClassifyViewModel`, `FolderSearchViewModel`,
`FolderPickerViewModel`), which now read `FolderMatchMode`/`MaxResults`
via `Settings.Load` so the new pane's settings actually take effect rather
than being cosmetic — this was the one unavoidable, narrowly-scoped touch
to already-shipped code.

---

## Step 10 — Polish, parity sign-off, packaging

- **Parity pass** against the §9 checklist in
  [legacy-overview.md](legacy-overview.md). Each row demoed on the
  32-bit target.
- **Reminder-to-front (§6d): DROPPED (confirmed 2026-06-07).** Purely
  cosmetic, hard-coded French reminder-count strings ("1 Rappel"… "15
  Rappels"), and implemented via `AppActivate` window-foregrounding — the
  exact MSForms-era focus-juggling hack already excluded elsewhere in this
  roadmap as obsoleted by WPF (see "Explicitly NOT reimplemented"). Not
  reconsidered; item 14 of the §9 checklist is marked dropped with this
  rationale.
- **Packaging: WiX per-user MSI — done, dev-machine-validated
  (2026-06-07).** `installer/Package.wxs` (built via WiX Toolset v7,
  installed as a per-user `dotnet tool`; `Build-Installer.ps1` wraps
  `wix build`, reading `release.config.json` as the same single source
  of truth `/make-release` and `/reload-addin` use). `Scope="perUser"`
  → installs to `%LocalAppData%\RBLclass\AddIn`, writes only `HKCU`, no
  elevation — declaratively replicating every registry write
  `Install-RblClassAddIn.ps1` makes: both CLSID subtrees
  (`CLSID\{guid}` + `Wow6432Node\CLSID\{guid}`) for both the add-in and
  task-pane-host classes, both ProgIds, the `.NET`/ActiveX-Control
  category markers, and the `Outlook\Addins\RBLclass.AddIn` key
  (`LoadBehavior=3`, `CommandLineSafe=1`). Test-installed on the dev
  machine with `msiexec /i ... /qn` over the existing PowerShell-kit
  install — Outlook loaded the add-in entirely from the MSI-written
  registration, ribbon tab confirmed working ("Works as expected").
  One harmless quirk found and verified inert: WiX's `[#FileId]`
  `CodeBase` token resolves with `/`-separators in the native CLSID
  subtree but `\`-separators in the `Wow6432Node` mirror — both parse
  to an identical `System.Uri.LocalPath`/`AbsoluteUri` (the same
  normalisation Fusion/`mscoree` applies), so both load correctly; not
  worth fighting WiX's resolver over a cosmetic difference. **Signing
  optional** per Step 0 (unsigned, as already validated for the POC
  kit). Per CLAUDE.md, the POC PowerShell installer remains the
  primary, validated path — both coexist — until the MSI is also
  installed and smoke-tested on the real 32-bit target workstation
  (the `Wow6432Node` subtree is the one the MSI write to that the dev
  machine's 64-bit Outlook never reads).
- **Regression matrix** for Semi-Annual channel updates (install, index,
  search, open, classify, send-guard, attachment removal, sent triage,
  settings).

---

## Explicitly NOT reimplemented

- **`RIC_rdv_vdev`** meeting-room finder — depends on Exchange/GAL
  FreeBusy, which the target excludes (PST is the source of truth).
- **`tests.bas` / `test_TW.frm`** — scratch code (its Access-index
  prototype is honoured by our SQLite folder table, but the code itself
  is not ported).
- The MSForms-specific UI hacks (`pleaseWait` progress shim, P/Invoke
  window placement, `AppActivate` focus juggling) — obsoleted by
  background threading + WPF.

---

## Summary of deviations — **approved 2026-06-06**

These were reviewed and approved by the user; several correct inaccurate
statements in the original CLAUDE.md / ROADMAP.md, which have been
reconciled accordingly.

| # | Legacy behaviour | Approved deviation | Why |
|---|---|---|---|
| 1 | Mail index implied by ROADMAP Phase 2 / FTS5 | **Defer** the mail FTS5 index to Phase 5; parity needs only a **folder** index | Legacy indexes folders only; shorter, lower-risk path |
| 2 | Folder search via in-memory pipe-strings | Substring match (in-memory over the cached tree, or SQL `LIKE`), **not FTS5**. **Index still persists in SQLite** — this changes the matching algorithm only, not storage | FTS5 tokenising breaks `"PJ"`→`ProjetJuridique` substring semantics |
| 3 | Settings in positional `options.txt` | SQLite `Settings` table | Mandated by stack; removes brittle line-positional parsing |
| 4 | Internal = address contains `CORPMAIL` | `AddressEntry` type + configurable internal-domain list | Robust, works without Exchange |
| 5 | Latest-reply split via HTML separator string | Prefer OM markers; else **configurable** marker+keywords | Theme/version/locale-proof |
| 6 | Task sentinel `TaskCompletedDate=01/01/4501` | Proper OM status / null check | Locale-formatted magic date is fragile |
| 7 | Conversation widening by full-folder iteration | `Restrict`/`Find`/`GetConversation` | Performance + correctness |
| 8 | Hard-coded FR/EN folder & window strings | Configurable exclusions; drop `AppActivate`/reminder hacks | De-localise; WPF handles focus/DPI |
| 9 | Authenticode signing as a gate (current ROADMAP) | Signing **optional**, not a functional gate | POC proved unsigned per-user install loads & passes EDR |
| 10 | Floating MSForms dialogs | Docked WPF Custom Task Pane (modeless window optional) | Matches chosen UI stack |

None of these change the **architecture** or the **stack** — they keep
us inside the COM-add-in / Core / Adapter / SQLite / WPF design. They
only replace locale- and version-fragile *implementation details* with
robust equivalents, and they re-sequence work so a usable tool ships
sooner. **Approved** — this plan is the baseline for implementation.
</content>
