# Legacy RBLclass — feature memo

This memo documents the legacy **RBLclass** Outlook VBA project, as
reverse-engineered from its source. It is the functional reference for
the reimplementation — written so it stands on its own, because the
legacy source itself is **not committed** to the repo (the
`RBLclass_legacy/` working-tree folder is gitignored and kept as a
local-only reference; file names cited below are from that local copy).
The companion document
[legacy-reimplementation-roadmap.md](legacy-reimplementation-roadmap.md)
maps each feature below onto our target COM-add-in architecture.

> Origin context: the macro was written for a French-language Outlook at
> AREVA NP (note the hard-coded "AREVA NP official colors", French folder
> names, French window titles, and the `CORPMAIL` Exchange marker). Many
> behaviours are locale-coupled; these are flagged as **fragility points**
> below because they must be de-coupled, not copied verbatim.

## 1. Purpose in one sentence

RBLclass lets a user **classify (file) emails into a deep folder tree
spread across several `.pst` archives**, by typing a few keywords to
locate the destination folder instead of dragging through the folder
pane — plus a handful of send-time and sent-item productivity guards.

## 2. Runtime shape

It is a classic **Outlook VBA project** (`ThisOutlookSession` +
standard modules + MSForms UserForms), not a COM add-in. Everything
lives in one Outlook profile and runs in-process on the UI thread.

| VBA artifact | Role |
|---|---|
| `ThisOutlookSession.cls` | Event sink: `Application_Startup`, `Application_ItemSend`, `colSentItems_ItemAdd`, `Application_Reminder`, `explo_SelectionChange` |
| `RBLclass_index.bas` | Builds the in-memory folder index at startup |
| `RBLclass_keywords.bas` | Keyword search over the folder index + result display |
| `RBLclass_folders.bas` | "Open a folder" action + sub-folder creation |
| `RBLclass_items.bas` | "Classify mail(s)" action + attachment removal |
| `RBLclass_setttings.bas` | Global settings/flags (note the triple-t typo) |
| `PSTArchive.cls`, `FolderInPST.cls` | The folder-tree domain model |
| `RBLclass_*.frm` | UserForms (class, open, settings, sent, sendExt, quickOpen, pleaseWait) |

## 3. Domain model — the folder index

Two class modules hold a pure in-memory model of the folder tree
(**folders only — mails are never indexed**):

- **`FolderInPST`** — one tree node:
  - `name`, `father`, `sons` (collection of `FolderInPST`), `nodeType`
    (`"root"` / `"inter"` / `"leaf"`).
  - `up` — the pipe-delimited path of the node and all its ancestors,
    e.g. `Archive2024|Projects|ContractX`.
  - `down` — `up` concatenated with the `down` of every descendant.
    So `down` is effectively *"every folder name appearing anywhere in
    this subtree, including the path that reaches it"*.
  - `upContainsKW(kws)` / `downContainsKW(kws)` — return true iff **all**
    space-separated keywords appear (substring, **case-insensitive**,
    **whitespace-stripped**) in `up` / `down` respectively.
- **`PSTArchive`** — `name` + `masterfold` (the root `FolderInPST`).
  The comment says it is "quite useless" — it is just a named handle
  to a tree root.

`indexationData` is a global `Collection` of `PSTArchive`, one per open
store.

### How the index is built (`RBLclass_index.bas`)

`indexFolders` runs at startup. It iterates the MAPI namespace's
top-level `Folders` (= open stores), **skipping** any store whose name
contains `"Dossiers publics"`, `"Public Folders"`, or `"G-FRA"`. For
each kept store, `archiveExplo` → `inDepthExplo` recurses the entire
subtree, skipping **Deleted Items** (`"Éléments supprimés"` /
`"Deleted Items"`) and public folders, building the `up`/`down` strings
on the way back up. A non-modal `RBLclass_pleaseWait` form shows the
store name as it works.

The index is **rebuilt** (for one store) whenever a sub-folder is
created. There is no incremental update and no persistence — the whole
tree is re-walked on every Outlook start.

**Fragility points:** store/folder exclusion is by hard-coded
French/English substrings (DE users break — the code says so);
the model holds live folder-name strings only, no `EntryID`/`StoreID`
stability.

## 4. Keyword search (`RBLclass_keywords.bas` + `FolderInPST`)

This is the heart of the tool. Given a space-separated keyword string:

- `createDico` runs `kwSearch` over every archive's root.
- `kwSearch` is a **pruned DFS**:
  - If not all keywords are present in the node's `down` subtree → stop
    (prune the whole branch).
  - Else if all keywords are present in the node's `up` ancestor-path →
    this folder **and all its descendants match**. Emit a result:
    - leaf → emit its `up` path;
    - non-leaf → either emit `up & "|[...]"` (collapsed summary) or,
      if the **All results** option is on, expand every descendant
      (`addDescendantsToDico`).
  - Else (keywords are scattered between ancestors and descendants) →
    recurse into the children.
- `displayDicos` merges the per-field result dictionaries, **aborts with
  a warning if the count exceeds `settingsMaxResults`**, otherwise
  inserts results into the list box **sorted alphabetically** and sizes
  the column to the longest entry.

Matching semantics worth preserving exactly: **AND across keywords**,
**substring (not token) match** (so `"PJ"` matches `ProjetJuridique`),
**accent/space-insensitive-ish** (spaces are stripped; case is ignored
via `vbTextCompare`). This substring/partial behaviour is the main
reason a naive FTS5 port would change results — see the roadmap.

## 5. Actions

### 5a. Open a folder — `RBLclass_open` + `RBLclass_folders.openFolder`

Type keywords → live-filtered folder list → pick one → Outlook Explorer
navigates to it. Options: **new window** (`MAPIFolder.Display`) vs same
window (`Explorer.CurrentFolder = …`), and **All results**.
`setTargetFolder` re-resolves the chosen pipe-path back to a live
`MAPIFolder` by walking `Namespace.Folders(part)` step by step.
UX: type-ahead search, **Enter** = select first hit + proceed,
**Ctrl** = toggle All results, double-click = proceed.

### 5b. Classify mail(s) — `RBLclass_class` + `RBLclass_items.classMail`

The flagship action. Triggered on the current Explorer selection
(`classMail_noarg` wraps the selection into a collection — the core
`classMail` only ever takes a collection, so it can also be driven by
the sent-item prompt).

Flow:
1. Show the classify form. **Multi-select** result list → one or more
   destination folders. Per-action toggles seeded from settings:
   **include whole conversation**, **keep a copy** (= don't delete
   originals), **All results**, **remove attachments**. Sub-folder
   creation is available inline.
2. **Conversation widening** (if checked): scan Inbox and Sent Items for
   any item sharing a selected item's `ConversationID` and add it to the
   set (keyed by `EntryID` to dedupe). Skips `IPM.Note.SMIME`.
3. **Task-completion guard:** if any item is flagged as a task and not
   completed (detected via `IsMarkedAsTask` and the sentinel
   `TaskCompletedDate = 01/01/4501`), prompt Yes/No/Cancel to mark such
   tasks completed in the destination.
4. **Attachment removal** (if checked): strip attachments first
   (with confirmation).
5. **The move:** for **each destination × each item**, `Copy` the item,
   optionally set the copy's `TaskCompletedDate`, then `Move` the copy
   to the destination. Filing to N folders therefore produces N copies.
6. **Delete originals** unless *keep a copy* is set.
7. Remember `lastProcessedFolder`; pop the **Quick Open** window.

**Fragility points:** the `01/01/4501` task sentinel and conversation
scan-by-full-iteration are slow and locale/version-coupled; the
copy-per-destination semantics must be preserved deliberately.

### 5c. Create a sub-folder — `RBLclass_folders.createSubFold`

From either action, type a name → `MAPIFolder.Folders.Add` under the
selected folder → **re-index that store** so it appears in subsequent
searches.

### 5d. Quick Open popup — `RBLclass_quickOpen`

A tiny borderless window shown top-right after a classify (only if the
**QuickOpen** setting is on). One button jumps the Explorer to
`lastProcessedFolder`. It auto-hides on the next `SelectionChange`
(managed by the `flagQuickOpen` toggle in `explo_SelectionChange`). Its
screen position is computed from `GetDeviceCaps` P/Invoke (DPI-aware
top-right placement).

### 5e. Attachment removal — `RBLclass_items.removeAttachments`

Remove **all** attachments from the selected mails (confirmation first),
`Save` each item. `removeAttachmentsFromSelection` is the standalone
entry point; `removeAttachments(coll)` is reused inside classify.

## 6. Send-time and sent-item guards (`ThisOutlookSession`)

### 6a. External-recipient warning — `Application_ItemSend`

If the **SendExt** setting is on, for every recipient whose
`Address` does **not** contain `"CORPMAIL"` (the marker for an internal
Exchange recipient), a modal warning (`RBLclass_sendExt`) is shown;
the user confirms or cancels the send. A pile of `AppActivate` calls
with French window titles tries to restore focus afterwards.
**Fragility point:** internal-vs-external is decided by the `CORPMAIL`
substring — brittle and environment-specific.

### 6b. Forgotten-attachment warning — `Application_ItemSend`

Only for `olMail`. It isolates *the latest reply text* via an HTML
heuristic — everything between `<body` and the first Outlook reply
separator (`solid #B5C4DF 1.0pt`) — then, if there are **no
attachments** but the text contains any of `attach`, `enclos`, `joint`,
`PJ`, it prompts before sending.
**Fragility point:** the separator string is Outlook-theme/version
specific; the keyword list is French/English only.

### 6c. Sent-item triage — `colSentItems_ItemAdd`

If the **UFsent** setting is on, when an item lands in **Sent Items** a
prompt (`RBLclass_sent`) offers: **Class**, **Delete**, **move to
Inbox**, or **Leave**, with an option to apply to the **whole
conversation** (Inbox + Sent). The event is detached during processing
(`Set colSentItems = Nothing` then re-`initialize_colSentItems` in a
`finally` label) to avoid re-entrancy when a copy transits Sent Items.

### 6d. Reminder-to-front — `Application_Reminder`

Brings the reminder window to the foreground via `AppActivate` with
hard-coded French titles (`"1 Rappel"` … `"15 Rappels"`). Cosmetic and
entirely locale-coupled.

## 7. Settings (`RBLclass_settings` + `RBLclass_setttings.bas`)

Nine settings persisted as **plain text, one value per line**, in
`%APPDATA%\Microsoft\Outlook\options.txt`:

| Global | Meaning |
|---|---|
| `settingsUFsent` | Show the sent-item triage prompt |
| `settingsSendExt` | Show the external-recipient warning |
| `settingsAllResults` | Expand all matching folders (vs `[...]` collapse) |
| `settingsWholeConversation` | Default "include whole conversation" on classify |
| `settingsdeleteProcessedElts` | Delete originals after classify (vs keep a copy) |
| `settingsNewWindow` | Open folders in a new window |
| `settingsRemoveAttachments` | Default "remove attachments" on classify |
| `settingsQuickOpen` | Show the Quick Open popup after classify |
| `settingsMaxResults` | Cap on displayed folder results (default 100, ±20 spinner) |

First-run / missing-file logic writes defaults and **forces the settings
form to stay on top until Saved** (`flagRBLclassSettings`, enforced in
`explo_SelectionChange`). Reads are positional and brittle: a missing
line throws and triggers the recreate-defaults path.

## 8. Out of scope / noise in the legacy project

These ship in the same VBA project but are **not** part of RBLclass's
classification purpose and should **not** be reimplemented as-is:

- **`RIC_rdv_vdev.bas` + `RIC_rdv_vdev_userform.frm`** — a meeting-room
  finder that queries the **Global Address List `FreeBusy`** over
  Exchange. It depends on Exchange/GAL connectivity, which our target
  explicitly excludes (PST is the source of truth). Drop it.
- **`tests.bas`** — scratch/experimental code. Notably it contains an
  **Access/SQL prototype** that indexes folders into `folders` +
  `parentChild` tables keyed by `EntryID` with a `fullPath` column.
  This is *informative*: the original author already reached for a
  relational folder index — which is exactly the SQLite direction our
  target takes.
- **`test_TW.frm`** — a TreeView experiment.
- `RBLclass_pleaseWait` — a progress shim only needed because VBA blocks
  the UI thread; our background-indexing design removes the need.

## 9. Feature checklist (for parity tracking)

1. Multi-PST folder-tree indexing (folders only), with exclusions.
2. Keyword folder search: AND, substring, case/space-insensitive, with
   collapse/expand and a result cap.
3. Open-a-folder navigation (same/new window).
4. Inline sub-folder creation + re-index.
5. Classify selection to one or many folders (copy-per-destination,
   optional delete-originals).
6. Conversation widening across Inbox + Sent.
7. Task-completion guard on classify.
8. Attachment removal (standalone + during classify).
9. Quick Open "jump to last filed folder" popup.
10. External-recipient send warning.
11. Forgotten-attachment send warning.
12. Sent-item triage prompt (class/delete/inbox/leave, whole
    conversation).
13. Settings persistence (9 options) + first-run defaults.
14. *(Cosmetic, optional)* reminder-to-front.
</content>
</invoke>
