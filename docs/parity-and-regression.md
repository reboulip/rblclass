# Parity sign-off & regression matrix

Step 10 deliverable. Two related checklists:

1. **Parity checklist** — one-time sign-off that every row of the legacy
   [§9 feature checklist](legacy-overview.md#9-feature-checklist-for-parity-tracking)
   has either landed in the new add-in or been consciously dropped (with
   a recorded rationale — nothing fell through by omission).
2. **Regression matrix** — a repeatable smoke test to re-run after any
   Outlook **Semi-Annual Enterprise Channel** update, since that's the
   only update cadence the deployment target tracks and the one thing
   most likely to silently break a COM add-in between releases.

Both were built up incrementally — every step in
[legacy-reimplementation-roadmap.md](legacy-reimplementation-roadmap.md)
ended in a live demo the user verified ("Works as expected") on the
64-bit ARM64 dev machine. The "32-bit target" column below is the one
piece that can only be checked off on the real workstation
(see CLAUDE.md "Dev Outlook" vs "Target Outlook").

## 1. Parity checklist

| # | Legacy feature (§9) | Landed in | Dev-machine (64-bit) | 32-bit target |
|---|---|---|---|---|
| 1 | Multi-PST folder-tree indexing, with exclusions | Step 1 — `FolderIndexService`, `FolderExclusionPolicy` | ✅ verified live | ☐ |
| 2 | Keyword folder search (AND, substring, case/space-insensitive, collapse/expand, result cap) | Step 2 — `FolderSearchService`; match-mode and cap now user settings (Step 9, [[folder-search-match-mode]]) | ✅ verified live | ☐ |
| 3 | Open-a-folder navigation (same/new window) | Step 3 — task pane + `NavigateTo` | ✅ verified live | ☐ |
| 4 | Inline sub-folder creation + targeted re-index | Step 4 — `IMailStore.CreateSubfolder` | ✅ verified live | ☐ |
| 5 | Classify selection to one or many folders (copy-per-destination, optional delete-originals) | Step 4 — `IClassifier.Classify`, `ClassifyViewModel` | ✅ verified live | ☐ |
| 6 | Conversation widening across Inbox + Sent | Step 6 — `Items.Restrict`/DASL, replacing the legacy O(n) full-folder scan | ✅ verified live | ☐ |
| 7 | Task-completion guard on classify | Step 6 — `ConfirmMarkTasksComplete` prompt | ✅ verified live | ☐ |
| 8 | Attachment removal (standalone + during classify) | Step 4 — `IClassifier` toggle + standalone "Remove attachments" ribbon button; filed-copy-only semantics ([[classify-attachments-on-copy]]) | ✅ verified live | ☐ |
| 9 | Quick Open "jump to last filed folder" popup | **DROPPED** — Step 5. A pane-local infobar version was built and tried live; didn't land well. Not re-attempted as specified ([[step5-quick-open-dropped]]) | n/a | n/a |
| 10 | External-recipient send warning | Step 7 — `ExternalRecipientGuard`, configurable internal-domains list (now in Settings) | ✅ verified live | ☐ |
| 11 | Forgotten-attachment send warning | Step 7 — `ForgottenAttachmentGuard`, configurable keyword list (now in Settings) | ✅ verified live | ☐ |
| 12 | Sent-item triage prompt (class/delete/inbox/leave, whole conversation) | Step 8 — `SentItemTriageWindow`/`ViewModel`, re-entrancy guard | ✅ verified live | ☐ |
| 13 | Settings persistence (9 options) + first-run defaults | Step 9 — `Settings`/`SettingsViewModel`/`SettingsWindow`; SQLite `Settings` table replaces `options.txt` (deviation #3); QuickOpen dropped, three new options added (`FolderMatchMode`, `InternalDomains`, `ForgottenAttachmentKeywords`) — eleven total | ✅ verified live | ☐ |
| 14 | *(Cosmetic, optional)* reminder-to-front | **DROPPED** — Step 10 (confirmed 2026-06-07). Hard-coded French reminder-count strings, implemented via `AppActivate` — the exact MSForms focus-juggling hack already excluded elsewhere in the roadmap as obsoleted by WPF | n/a | n/a |

**Sign-off note:** every portable row (1–8, 10–13) has been demoed and
confirmed working on the dev machine across Steps 1–9; the two dropped
rows (9, 14) have a recorded decision and rationale. The 32-bit-target
column is the one section of this table that needs a session on the real
workstation to close out — see CLAUDE.md's bitness/EDR validation notes.

## 2. Regression matrix (re-run after each Semi-Annual Channel update)

Outlook's Semi-Annual Enterprise Channel is the one thing on the
deployment target that changes underneath the add-in without any action
on our part. COM add-ins can be silently disabled (`LoadBehavior` flipped
3→2) by an unhandled exception, and channel updates have been known to
shuffle ribbon ids, ribbon imageMso names, or COM threading behaviour.
Run this matrix after every channel update lands on the target — and
before cutting a release that follows one.

| Area | What to check | Pass criteria |
|---|---|---|
| **Install** | Run the install kit (`/make-release` output) on a clean profile, **and/or** `installer\bin\*.msi` (`installer\Build-Installer.ps1` → `msiexec /i`, dev-machine-validated 2026-06-07 — see roadmap Step 10) | Add-in loads (`LoadBehavior` stays 3); RBLclass ribbon tab appears; no startup error dialog |
| **Index** | First run after install/profile change | Background walk completes; "Index status" shows expected store/folder counts; <500ms contribution to Outlook startup (CLAUDE.md target) |
| **Search** | Type keywords in the folder-search task pane | Results appear within ~100ms for the indexed set (CLAUDE.md target); collapse/expand and match-mode settings are honoured |
| **Open** | Pick a search result, try both same-window and new-window | Outlook navigates to the right folder in the right window |
| **Classify** | Select mail(s), classify to 1 and to several folders, with keep-copy on and off, with attachments removed and kept, with conversation widening on and off | Copies/moves land correctly; originals handled per toggle; task-completion prompt appears when expected; perceived latency <300ms |
| **Send-guard** | Send to an external recipient; send without an attachment when the body mentions one | Both warnings fire (or don't, per settings) and can be acknowledged or cancelled |
| **Attachment removal** | Standalone "Remove attachments" on a multi-selection | Confirmation appears; attachments stripped from the selected mails only |
| **Sent triage** | Send a mail and let it land in Sent Items | Prompt appears (if enabled); each of Class/Delete/Move-to-Inbox/Leave works; no re-entrant re-prompt on the resulting copy/move |
| **Settings** | Open the Settings dialog, toggle a few options, close, reopen, restart Outlook | Every option persists across reopen and restart; guards/actions reflect the new values immediately (live-apply) |

Any failure here means: check Serilog's rolling file sink first
(`%LocalAppData%\RBLclass\logs\`, see CLAUDE.md), then the registry
`LoadBehavior` value before assuming a code regression — a channel update
disabling the add-in looks identical to a crash from the user's side.
