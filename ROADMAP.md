# RBLclass roadmap

**Closed history** — Phases 0–7 and every shipped version (**v2.0.0.0**
through **v2.12.0.0**) — lives in
[docs/roadmap-history.md](docs/roadmap-history.md). This file holds **where we
are**, the **open carried-over items**, the next sprint milestone once one is
planned, and the **forward-looking phases**.

Reference docs from the legacy analysis:
[docs/legacy-overview.md](docs/legacy-overview.md),
[docs/legacy-reimplementation-roadmap.md](docs/legacy-reimplementation-roadmap.md),
[docs/parity-and-regression.md](docs/parity-and-regression.md).

## Where we are (2026-11-21)

The rewritten add-in reached full legacy parity as **v2.0.0.0**, went through
the pilot and general availability (Phases 6–7), and has since shipped
feedback waves up to **v2.12.0.0** (experimental re-encrypted filing of
protected mail, 2026-11-21 —
[release](https://github.com/reboulip/rblclass/releases/tag/v2.12.0.0)). All
of them are recorded in [history](docs/roadmap-history.md). The next sprint
milestone, **v2.13.0.0** (pane UI study and redesign, search depth, PDF
export, phishing warnings), was planned on 2026-11-21 from a `/gh-issues`
triage.

Open work: the v2.13.0.0 sprint below, plus the two real-target items,
encrypted (Stormshield) mail [#62] and the external-sender banner on
reply/forward [#46]. Phases 8–9 are forward-looking.

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

Both items depend on mail this dev machine never receives (Stormshield/S-MIME,
the external-sender banner), so they are checked on the real 32-bit target and
stay out of sprint scope. The full specs and dev-side results are in the
archived version sections of [history](docs/roadmap-history.md).

- [ ] **Encrypted (Stormshield) mail: real-target verification** [#62].
      Covers v2.11.0.0 A3 (protected-mail guard and protection trace),
      v2.11.0.0 C2 (body images as removable attachments) and v2.12.0.0 A8
      (experimental re-encrypted filing). On a target workstation:
      1. **Step 0, first:** run `tools/Inspect-SmimeMsg.ps1` on a real
         Stormshield `.msg`. It must show a standard CMS envelope; record the
         key transport. RSA-OAEP keeps `ReencryptKeyTransport` = `Oaep`;
         PKCS#1 v1.5 means switching that default to `Pkcs1`; a non-standard
         envelope shelves the re-encryption feature.
      2. **Standard path, card inserted**, Diagnostic logging on: classify one
         protected mail per path (move, copy, safety copy, keep a copy) with
         "remove attachments" on. The attachments must stay, and the log must
         show `protected mail detected by <Reason>` (expected:
         `StoredMessageClass`). A `Protected mail filed by … no longer reads
         as protected` warning means Outlook filed it decrypted (a known
         object-model limit). Read the class and flags from the
         `Protection trace` lines and decide whether `PR_SECURITY_FLAGS` or
         another class prefix should drive the check (today: class-based,
         flags logged only).
      3. **Card removed:** classify still files the mail; if nothing happens,
         send the `GetSelectedItems` warning line from the log.
      4. **Body images option on** (v2.11.0.0 C2): a protected mail with body
         images keeps its images and its body untouched.
      5. **Experimental re-encryption on** (Settings → Experimental): a
         classified Stormshield mail is filed re-encrypted (stored
         `IPM.Note.SMIME`) and opens decrypted in Outlook with all its parts;
         the external banner is stripped from the rebuilt body; with the card
         removed, the mail is filed through the standard path.
      6. **Regression pass on ordinary mail:** copy, safety copy, mark as read,
         attachment removal, conversation widening.
      Close #62 once 1–6 pass, or once the feature is shelved at step 0.

- [ ] **External-sender banner not stripped on reply/forward** [#46].
      v2.9.0.0 B1 applies the strip at draft creation (`Inspector.Activate`
      rather than `NewInspector`, plus `Explorer.InlineResponse` for inline
      replies), but it is **confirmed still broken on the real target
      (2026-07-28)**: detection works (Settings → "Diagnose selected mail"
      reports a match, and the filed-copy strip works), the reply/forward
      strip does not. v2.10.0.0 D1 added Settings → "Save diagnostic file",
      which writes the body with the matched banner highlighted to
      `%LocalAppData%\RBLclass\diagnostics\`. Next steps: on the target, check
      that action (file written, match highlighted, folder created on first
      use), get a diagnostic file for a reply/forward draft, root-cause why
      the strip fails there, then fix and re-verify on the target.

---

## v2.13.0.0 — feedback wave: pane UI study and redesign, search depth, PDF export, phishing warnings

Planned 2026-11-21 from the open GitHub issues (`/gh-issues` triage). Each
item embeds the implementation decisions made with the user — treat the item
text as the specification.

**Sprint order (user decision).** The sprint opens with the two studies,
**A1** and **D1**, and stops after them for user validation. Their validated
outcomes are then written into **A2** and **D3** (which are gated on them),
committed as a `ROADMAP.md` update, and only then implemented. B1, C1 and D2
do not depend on the studies and run in the waves after that validation.

### A — Pane UI

- [x] **A1.** Critical study of the whole pane UI, with #67 as its central
      constraint. Docs only: no code change in this item. Write
      `docs/ui-review.md` (internal doc, not the website) containing:
      1. **Inventory**: every element of `MainPaneView.xaml` (rows 0–7: the
         update and pinned-mail banners, the header icon buttons and index
         status dot, the query box, the results list with its per-row 👁/⤢
         buttons and three path display modes, the destination chips, the
         Options toggle and panel, Classify/Undo, the search feedback and
         journal), both layout modes (vertical, and the horizontal mode of
         v2.9.0.0 A2), and the modals (Settings, attachment disposition,
         Save-as picker, sent-item triage, name prompt, meeting hint).
      2. **The #67 problem, measured**: on new HP laptops (smaller screens)
         the pane cannot be narrowed without clipping. The header is a
         right-docked `DockPanel` (`LastChildFill="False"`), so the leftmost
         items (the `?` button, the index status dot) are pushed out of view,
         and the `Classement auto` text button takes a large share of the
         width. Widening the pane instead makes Outlook drop the mail list's
         column view. Measure the minimum pane width at which nothing is
         clipped today, in EN/FR/DE, and set a **target minimum width**
         (starting proposal: 260 DIP) at which everything stays usable.
      3. **Critique against named best practices**: Windows/Fluent design
         guidance on visual hierarchy, a single prominent primary action,
         progressive disclosure, icon-only buttons (tooltip plus accessible
         name, icon consistency), click-target size, spacing and alignment
         grid, text truncation, keyboard flow (v2.11.0.0 B3 `/` and focus
         return, B4 focus ring), and WCAG AA contrast (4.5:1) in light and
         dark themes. Each finding names the element, the guideline, and the
         impact.
      4. **Font scaling**: assess the issue's suggestion (font size
         proportional to the pane width) against responsive layout
         (wrapping, overflow menu, icon-only modes), and give a
         recommendation with its rationale.
      5. **At least two layout proposals**, each with mockups at the target
         minimum width and at a wide width (ASCII or embedded images),
         stating what moves, what is merged, and what is dropped.
      6. **A prioritized change list** (each change tagged quick win or
         structural) and a **proposed A2 specification**.

      Constraints the proposals must respect: CLAUDE.md "never resize the CTP
      after creation" ([[never-resize-ctp-after-creation]]); WPF on .NET
      Framework 4.8 with the hand-rolled MVVM and no new dependency; every
      string in EN/FR/DE. Validate the study with the user through
      `AskUserQuestion`, then rewrite A2 from the validated outcome and commit
      that `ROADMAP.md` change before A2 starts. [#67]

- [ ] **A2.** Pane redesign: command bar with overflow, reflowing rows,
      three-column Top/Bottom layout (requires: A1). Files:
      - `Views/MainPaneView.xaml` and `.xaml.cs`;
      - `Views/Themes/Controls.xaml` and `Brushes.{Light,Dark}.xaml`;
      - `Resources/Strings{,.fr,.de}.resx`;
      - `RblClassTaskPaneHost.cs` (font mapping);
      - `RblClassAddIn.cs` (`ApplyDockSize` only, at creation);
      - `ViewModels/MainPaneViewModel.cs` (view-state properties only);
      - `Views/SentItemTriageWindow.xaml` and `SettingsWindow.xaml` (labels).

      No new dependency, no Core change.

      1. **Width contract.** At a view width of **200 DIP and above** (184 DIP
         column), no control is clipped and no label breaks mid-word, in
         EN/FR/DE, light and dark. Below 200 DIP, rows keep reflowing: no
         right-docked `DockPanel` chains anywhere in the pane. Set
         `d:DesignWidth="200"`.
      2. **Font and icons.** Set `FontFamily="Segoe UI" FontSize="12"` on the
         `MainPaneView` root. In `RblClassTaskPaneHost`, call
         `_elementHost.PropertyMap.Remove("Font")` right after creating the
         `ElementHost`. Secondary text (feedback, journal, banner subject) uses
         11 DIP and `MutedText`. Icons: `FontFamily="Segoe MDL2 Assets"`,
         `FontSize="16"`:

         | Command | Glyph |
         |---|---|
         | Save | `&#xE74E;` |
         | Attach | `&#xE723;` |
         | More | `&#xE712;` |
         | Refresh | `&#xE72C;` |
         | Settings | `&#xE713;` |
         | Help | `&#xE897;` |
         | Undo | `&#xE7A7;` |
         | Open folder (eye, keeps v2.9.0.0 A1) | `&#xE890;` |
         | Path display | `&#xE740;` |
         | Remove chip | `&#xE711;` |
         | Clear all | `&#xE8E6;` |
         | Chevrons | `&#xE76C;` / `&#xE70D;` |

         The easter-egg pig image is kept.
      3. **Command bar** (replaces `:88-144`). A `Grid` with columns `*, Auto,
         Auto, Auto`:
         - column 0: the Auto-class button, `HorizontalAlignment="Left"`,
           `Padding="8,4"`. Its label is `TextWrapping="WrapWithOverflow"` with no
           `MaxWidth`, and `AutomationProperties.Name` is
           `MainPane_AutoClass_Button`;
         - columns 1–3: Save as, Manage attachments and More, each 28×28 with
           `Margin="4,0,0,0"`, keeping their tooltips.

         More opens a `ContextMenu` (placement bottom, right-aligned) with Update
         folders, Settings…, Online documentation, a separator and a
         non-interactive line bound to `IndexStatusToolTip`. Each menu item calls
         the existing view-model method, then `ReturnFocusToSearch()`, exactly
         like today's click handlers. Declare children in visual order.
      4. **Search row.**
         - The existing `TextBox x:Name="QueryBox"` keeps every binding and
           handler, with `Padding="4,3,20,3"` and
           `AutomationProperties.Name=MainPane_QueryBox_Name`.
         - A placeholder `TextBlock` overlays it: `IsHitTestVisible="False"`,
           `MutedText`, trimmed, visible only while the box is empty.
         - An 8×8 status ellipse sits inside the box on the right
           (`Margin="0,0,8,0"`, same colour triggers, faint `Theme.Foreground`
           outline), with tooltip `IndexStatusToolTip` and name
           `MainPane_IndexStatus_Name`.
         - Under the box, a `MutedText` notice line shows `IndexStatusToolTip`
           only while the index is not ready, with
           `AutomationProperties.LiveSetting="Polite"`; raise
           `AutomationEvents.LiveRegionChanged` from its peer when the text
           changes. Add an `IsIndexReady` view-model property, raised with
           `IndexStatus`.
      5. **Results list.** Behaviour unchanged.
         - The item container style gets `AutomationProperties.Name="{Binding
           DisplayPath}"`; so does the row `CheckBox`.
         - Rows get `MinHeight="24"`.
         - The row buttons become 24×24 MDL2 icons with
           `MainPane_OpenFolder_Name` and `MainPane_ExpandPath_Name` (tooltips
           kept).
         - `v:PathTrim.RightReserve` matches the new button block.
      6. **Destinations.**
         - Remove the "Selected: N folders" header row (`:230-236`). Use
           `SelectionHeader` as the chips container's `AutomationProperties.Name`.
         - The chips container is a `Grid` with columns `*, Auto`: the chip
           `WrapPanel` in column 0, and a 24×24 clear-all icon (`&#xE8E6;`) in
           column 1, top-aligned, with the existing tooltip and name
           `MainPane_ClearAll_Name`.
         - Chip names get `MaxWidth="150"` and `TextTrimming="CharacterEllipsis"`
           (full-path tooltip kept).
         - The remove buttons are 20×20 with name
           `MainPane_RemoveDestination_ToolTip`.
      7. **Toggles row.** A `Grid` with columns `*, Auto`:
         - Options `ToggleButton` (`IsChecked` two-way to `IsOptionsExpanded`,
           `x:Name="OptionsToggle"`) on the left;
         - Journal `ToggleButton` (two-way to `IsStatusLogExpanded`, visible when
           `HasStatusEntries`, `x:Name="JournalToggle"`) on the right.

         Content is a chevron that turns via a trigger, plus `Pane_Options_Label`
         / `Pane_Journal_Label`. Add implicit `ToggleButton`, `ContextMenu`,
         `MenuItem` and `Separator` styles to `Controls.xaml`, using the theme
         brushes, with the checked state on `Theme.SelectionBackground` and the
         `Theme.FocusBorder` ring on `IsKeyboardFocused`, in both themes.
      8. **Options panel.** Each `CheckBox` gets `<TextBlock
         TextWrapping="Wrap" …/>` content. Keep `x:Name="FirstOptionCheckBox"`
         and the Tab-from-search behaviour.
      9. **Action row.** A `Grid` with columns `*, Auto`, a single instance for
         both modes:
         - Classify: `AccentButton`, stretched, `Padding="10,5"`, content a
           centred, wrapping `TextBlock` bound to `ClassifyButtonText`;
         - Undo: a 28×28 icon, `IsEnabled="{Binding CanUndo}"`, name
           `MainPane_Undo_Button`, tooltip `MainPane_Undo_ToolTip` with
           "(Ctrl+Z)" / "(Strg+Z)" appended.
      10. **Status.** The search feedback gets
          `AutomationProperties.LiveSetting="Polite"`, with `LiveRegionChanged`
          raised when `SearchFeedback` changes. The journal is unchanged.
      11. **Horizontal layout** (replaces v2.9.0.0 A2's bar).
          - `IsHorizontalLayout = ActualWidth >= 560 && ActualWidth >
            ActualHeight` (`MainPaneView.xaml.cs:32-35`).
          - When true, `DataTrigger` setters re-place the same elements into
            three columns: column 0 (`Width="300"`) holds the command bar, search,
            action row and status; column 1 (`Auto`, visible only while Options or
            Journal is open, in a vertical `ScrollViewer`) holds the Options panel
            and the journal; column 2 (`*`) holds the destinations above the
            results.
          - Delete the duplicated horizontal `StackPanel` (`:272-285`) and
            `CanShowJournalInVerticalMode`.
          - Check: at a 1200 × 163 DIP view, with one line of chips and Options
            open, at least 4 result rows (24 DIP each) are visible and nothing is
            clipped.
      12. **Contrast.**
          - Dark `Theme.AccentBackground` `#FF115EA3`, `AccentHover` `#FF0F548C`,
            `AccentPressed` `#FF0C3B5E`.
          - Update banner text `#FF201F1E` on `#FF8C00`.
          - A new `Theme.InputBorder` brush (light `#FF8A8886`, dark `#FFA19F9D`)
            for the `TextBox` and `ListBox` borders.
          - `SettingsWindow.xaml:100` and `:112` use `MutedText` instead of
            `#777`.
      13. **Banners and dialogs.**
          - The pinned-banner Delete button content wraps.
          - The Sent-items triage buttons get wrapping `TextBlock` content and
            the window gets `SizeToContent="Height"` (width 380, still not
            resizable).
      14. **Task-pane size (creation only).** `ApplyDockSize`, still called only
          from `EnsureTaskPane` before `Visible` and before the
          `DockPositionStateChange` subscription, sets:
          - Left/Right: `Width = Math.Max(360, ceil(229 × s) + 16)`;
          - Top/Bottom: `Height = Math.Max(300, ceil(207 × s))`;

          where `s = GetDpiForWindow(hwnd) / 96.0` and `hwnd` is the Outlook main
          window (`Process.GetCurrentProcess().MainWindowHandle`, as
          `DialogPlacement` uses; `GetDpiForWindow` requires Windows 10 1607,
          below the 1809 minimum). Log the requested size and `s`, and log the
          host's pixel size on its first `OnResize`. No size is ever set after
          creation.
      15. **Invariants (acceptance):**
          - the `/` jump;
          - focus return after every pane action, including the More menu items
            and the toggles;
          - Ctrl+Z, Enter / Ctrl+Enter, Ctrl tap, Tab to Options;
          - the focus ring on every focusable element in both themes;
          - a localized tooltip and `AutomationProperties.Name` on every
            icon-only button;
          - the CTP never resized after creation.
      16. **Strings.** Add these keys to all three resx files, remove
          `Pane_Options` and `Pane_Journal`, and keep `Status_SelectedFolders_*`
          (now an accessible name):

          | Key | EN | FR | DE |
          |---|---|---|---|
          | `MainPane_More_Name` | More actions | Plus d'actions | Weitere Aktionen |
          | `MainPane_More_ToolTip` | More: update folders, settings, online documentation | Plus : actualiser les dossiers, paramètres, documentation en ligne | Mehr: Ordner aktualisieren, Einstellungen, Online-Dokumentation |
          | `MainPane_UpdateFolders_MenuItem` | Update folders | Actualiser les dossiers | Ordner aktualisieren |
          | `MainPane_Settings_MenuItem` | Settings… | Paramètres… | Einstellungen… |
          | `MainPane_Help_MenuItem` | Online documentation | Documentation en ligne | Online-Dokumentation |
          | `MainPane_SaveAs_Name` | Save as .msg | Enregistrer en .msg | Als .msg speichern |
          | `MainPane_ManageAttachments_Name` | Manage attachments | Gérer les pièces jointes | Anlagen verwalten |
          | `MainPane_OpenFolder_Name` | Open folder | Ouvrir le dossier | Ordner öffnen |
          | `MainPane_ExpandPath_Name` | Change path display | Changer l'affichage du chemin | Pfadanzeige ändern |
          | `MainPane_ClearAll_Name` | Clear all destinations | Retirer toutes les destinations | Alle Ziele entfernen |
          | `MainPane_QueryBox_Name` | Search folders | Rechercher des dossiers | Ordner suchen |
          | `MainPane_QueryBox_Placeholder` | Search folders… | Rechercher un dossier… | Ordner suchen… |
          | `MainPane_IndexStatus_Name` | Folder index status | État de l'index des dossiers | Status des Ordnerindex |
          | `Pane_Options_Label` | Options | Options | Optionen |
          | `Pane_Journal_Label` | Journal | Journal | Journal |

          Also reword `IndexStatus_NotFound` to point at "⋯ > Update folders".
          This also fixes the EN wording "Refresh folders", which does not match
          the command's name.
      17. **Docs** (doc-keeper): update the walkthrough's "Pane icon buttons"
          section, the shortcuts note (Undo icon) and the pane-position note
          (sizes).
      18. **Verification.** A live check with the user at a 200 DIP view (about
          316 px of frame at 150 %) and at the default width, in EN/FR/DE, light
          and dark, docked Left/Right and Top/Bottom. Optionally run an offline
          layout check with the method in Appendix A of `docs/ui-review.md`.
      Validated by the user on 2026-11-21 (study: `docs/ui-review.md`): target
      200 DIP of view width, Proposal B, three-column Top/Bottom, Segoe UI 12 + MDL2
      icons, no font scaling. Acceptance criteria fixed earlier still apply. [#67]

### B — Folder search

- [ ] **B1.** Add a maximum search depth for the folder search. New setting
      `MaxSearchDepth` (`SettingsKeys`, parsed in `Settings` like
      `MaxResults`): an integer, default **0 = unlimited**, negative values
      clamp to 0. It sits in Settings → Search as a numeric field labelled
      "Maximum search depth (0 = unlimited)", translated FR/DE.
      **Levels:** the store (PST) root is level 0 and its direct sub-folders
      are level 1. Compute a folder's depth from the `ParentEntryId` chain
      (keyed by `(StoreId, EntryId)`), **not** by splitting `FullPath`,
      because folder names may contain `/`. Cache depths alongside the
      normalized-path cache in `FolderSearchService` and rebuild them on the
      same `CacheVersion` change. Add `MaxDepth` to `FolderSearchOptions`
      (0 = unlimited) and wire it from the settings like `MaxResults`.
      **Behaviour (user decision: deeper folders are excluded):** a folder
      deeper than N is not a match candidate. It can still appear only
      through the existing collapse, as part of a matching ancestor of level
      ≤ N shown as `path | [...]`. With All results on, folders deeper than
      N are not listed either, and a listed level-N folder that has
      sub-folders gets the `FolderSearchResult.CollapsedSuffix` so the user
      can tell its subtree was cut. `MaxResults` applies after the depth
      filter. Favourite filesystem folders (`FavoriteFolderSearchAdapter`,
      flat) are not affected. xUnit cases: 0 means unlimited; a deep-only
      match is excluded; a matching ancestor at level ≤ N still collapses
      over its deep subtree; with All results on, a level-N folder with
      children carries the suffix; the store root is level 0; a folder name
      containing `/` gets the right depth; the cap counts results after
      depth filtering. Update the walkthrough settings table. [#68]

### C — Save as

- [ ] **C1.** Add PDF export to the pane's 💾 Save-as button (v2.8.0.0
      A1, today `.msg` only through `OutlookMailStore.SaveMailAsFile`).
      **Setting** `SaveAsFormat`: `Msg` (default), `Pdf`, `Ask`, in Settings
      next to the other save/attachment options, labelled "Outlook message
      (.msg)", "PDF (.pdf)", "Ask each time" (translated FR/DE). With `Ask`,
      the Save-as picker window shows a two-option format choice (`.msg`
      preselected) above its OK button; with `Msg` or `Pdf` the picker looks
      as it does today. The `.msg` path is unchanged.
      **PDF conversion** (no new dependency, no Word PIA reference):
      1. On the Outlook UI thread, per mail: `mail.SaveAs(<temp>.mht,
         olMHTML)` into `%TEMP%\RBLclass\pdf\` (random file name). MHTML
         carries the header block (From, Sent, To, Cc, Subject, Attachments),
         the body and inline images.
      2. On a **dedicated STA background thread** (Word runs out of process;
         Outlook OM calls stay on the UI thread): create one Word instance per
         batch through late binding (`Type.GetTypeFromProgID("Word.Application")`
         + `Activator.CreateInstance`, called via `dynamic` or
         `InvokeMember`), `Visible = false`, `DisplayAlerts = 0`; for each
         file `Documents.Open(path, ConfirmConversions: false, ReadOnly:
         true, AddToRecentFiles: false, Visible: false)`,
         `ExportAsFixedFormat(<dest>.pdf, 17 /* wdExportFormatPDF */)`,
         `Close(0)`; `Quit(0)` in a `finally`, and `Marshal.ReleaseComObject`
         on every Word object.
      3. Always delete the temp `.mht` files, even on failure.

      The pane shows its busy overlay with progress ("Saving as PDF… n/N").
      File name: the same sanitized subject as `.msg`, with `.pdf`; on a
      collision use `AttachmentFilenameCollisionResolver` (suffix). No
      checksum dedupe, since PDF bytes differ from one run to the next. If
      Word is missing or fails to start, write nothing and show a localized
      status "PDF export needs Microsoft Word". Only `MailItem`s are saved,
      as for `.msg`.
      **Encrypted mail (user decision: ask):** before converting, partition
      the batch with the shared protected-mail check (`ProtectedMailPolicy`).
      If any mail is protected, show one confirmation for the whole batch:
      "N of the selected mails are encrypted. The PDF will contain them
      unencrypted." with **Continue** and **Skip encrypted mails** (default:
      Skip). Skipped mails are reported in the pane journal with their own
      `_One`/`_Other` status keys, separate from failures.
      **Core:** the format resolution (setting plus the Ask choice), the PDF
      file naming and the protected/plain partition live in a small
      `RBLclass.Core` planner with xUnit tests.
      **Docs:** walkthrough settings table; `docs/security.md` and
      `website/docs/security.md` (plaintext PDF of encrypted mail only after
      explicit confirmation; out-of-process Word automation; temp MHTML
      cleaned up); tech-stack (Word late-bound, optional at runtime).
      Real-target check: Word automation is not blocked by the EDR on the
      32-bit target, and the confirmation appears for a Stormshield mail
      (not verifiable on the dev machine). [#69]

### D — Phishing warnings on external-banner mail

The goal (user decision): the user is warned **systematically, with no
action of their own**, when they open a link or an attachment from a mail
carrying the organization's external-sender banner, because a user who
already has a doubt does not need the tool. A warning before the link opens
is preferred; immediately after is acceptable if "before" is not feasible.
Applies to every item: a mail "carries the banner" when
`ExternalBannerStripper.ContainsBanner(htmlBody, <learned signature>)` is
true (signature: `SettingsKeys.ExternalBannerSignature`). The body is read
**only when the event fires**, never on each selection change (performance,
and reading a Stormshield body prompts for the smart card). No warning when
no signature has been learned. No global keyboard or mouse hook, no API
patching, no modification of the stored mail.

- [x] **D1.** Study: intercepting link clicks in banner mail. The Outlook
      object model has no hyperlink-click event, in either the reading pane
      or the inspector (the reason #37 was deferred on 2026-07-03).
      Prototype each candidate on the dev machine with throwaway code
      (not committed, or kept under `tools/` if useful later) and assess it
      on: before or after the open, whether it can cancel, whether the URL is
      known, coverage (reading pane, opened inspector, inline reply), EDR
      risk on the target, S/MIME impact, and performance. Candidates:
      1. **UI Automation** on the reading pane and inspector: hyperlink
         `Invoke` events, or `AutomationElement.FromPoint` at click time,
         from a UIA client on a background MTA thread inside Outlook's
         process (check for re-entrancy and deadlocks).
      2. A **thread-scoped** `WH_MOUSE` hook on Outlook's UI thread only
         (`SetWindowsHookEx` with that thread id, not a global hook) plus a
         UIA hit test of the element under the cursor to read the link's URL
         before Outlook follows it, and whether swallowing the click cancels
         the navigation.
      3. An out-of-context `SetWinEventHook(EVENT_SYSTEM_FOREGROUND,
         WINEVENT_OUTOFCONTEXT)` noticing a browser taking the foreground
         right after the user acted on a banner mail (after the fact, URL
         unknown, browser title available).
      4. **Rewriting links in the stored body** to a local handler (custom
         HKCU protocol): rejected a priori (modifies received mail, breaks
         S/MIME signatures, adds an executable and a trust boundary). Record
         why in the study.
      5. Word events on the inspector's `WordEditor` (`WindowSelectionChange`)
         and whether they reach Outlook-hosted Word at all.

      Deliverable: `docs/phishing-link-study.md` with the comparison matrix,
      the prototype findings, and a recommendation in this order of
      preference: before the open, otherwise after it, otherwise not
      feasible. Validate it with the user through `AskUserQuestion`, then
      rewrite D3 from the validated outcome and commit that `ROADMAP.md`
      change before D3 starts. [#37]

- [ ] **D4.** Fix: the banner strip must never modify a received mail
      (found by D1, `docs/phishing-link-study.md` section 2.12). With *Strip
      banner from replies* on, `OnNewInspector` (`RblClassAddIn.cs`, ~754)
      subscribes `Inspector.Activate` for every inspector and its handler
      calls `_mailStore.StripBannerFromDraft(inspector.CurrentItem, sig)`.
      `StripBannerOnLiveMail` (`OutlookMailStore.cs`, ~1827-1851) rewrites
      `HTMLBody` and calls `Save()` without checking `mail.Sent`, so opening a
      received banner mail in its own window removes the banner from the
      stored mail for good (measured: body 1283 to 514 characters,
      `Sent=True`) and blinds D2/D3. Strip only drafts (`mail.Sent == false`),
      in `StripBannerFromDraft` itself so that every caller is covered; when
      the mail is sent, do nothing and log at Debug. Keep behaviour for
      reply/forward drafts unchanged. Add an adapter-level guard comment, and
      a Core-free test is not possible, so Claude verifies end to end: open a
      received fake banner mail (own window) with the setting on and check via
      COM that the stored body is unchanged and `Sent` is true, and that a
      reply draft is still stripped. No setting or string change. [#37]

- [ ] **D2.** Warn before an attachment from a banner mail opens. Hook the
      item events `BeforeAttachmentWriteToTempFile` and
      `BeforeAttachmentPreview` (both take `ref bool Cancel`) on (a) the
      single selected mail of the active explorer, re-hooked on the
      debounced `SelectionChange` (multi-selection or non-`MailItem`: no hook),
      and (b) each opened inspector's `CurrentItem` (hooked on
      `NewInspector`, released on `Inspector.Close`). Keep the item and
      event-source references in fields so the events keep firing, release
      the previous ones through `ComRef<T>` on every change, and unhook in
      `OnBeginShutdown`. Verify empirically and record in the item which
      events fire for double-click open, reading-pane preview and drag to the
      desktop. Saving an attachment (`BeforeAttachmentSave`) does not warn.
      When an event fires and the setting is on and the mail carries the
      banner, show a modal: title "Possible phishing", text "This attachment
      comes from an external email. It may be a phishing attempt.", the
      sender display name and real SMTP address (resolved as the existing
      external-recipient code does), the attachment file name, and buttons
      **Open anyway** / **Cancel** (Cancel is the default and `Esc`). Cancel
      sets `Cancel = true`. Top-level try/catch in each handler: on error,
      log and let the attachment open (fail open, as the send guard does).
      New setting `PhishingWarnAttachments`, default **on**, in Settings
      next to the external-banner options. **Core:** a
      `PhishingWarningPolicy` (setting, learned signature, HTML body →
      warn or not) and the message model, with xUnit tests. New strings in
      EN/FR/DE. Update the walkthrough and both `security.md` files. Dev
      verification: I run it end to end with a fake mail carrying a learned
      banner (UIA + COM, [[e2e-outlook-testing]]). Real banner mail is not
      verifiable on the dev machine: leave the checkbox open with
      `- shipped, pending tests` and re-verify on the real target. [#37]

- [ ] **D3.** Warn before a link from a banner mail opens (requires: D1,
      D4, and D2 for `PhishingWarningPolicy`/the warning window if D2 lands
      first; otherwise create them here). **Validated by the user on
      2026-11-21 from `docs/phishing-link-study.md`; section 5 of that study
      is the implementation spec and is authoritative (classes, hook
      procedures, hit test, sender resolution, strings table, Open-anyway
      logic, cleanup). Read it fully before implementing.** Summary of the
      validated decisions: warning **before** the open, systematic, banner
      mails only, URL known, cancellable; mechanism = two **thread-scoped**
      hooks on Outlook's UI thread (`WH_MOUSE` for clicks, `WH_GETMESSAGE`
      for Enter after Tab; never global, hMod NULL) plus repurposing the
      ribbon command `HyperlinkOpen` (`<commands>` block in `Ribbon.xml`,
      `.fr.xml`, `.de.xml`; callback `OnHyperlinkOpenRepurposed(IRibbonControl
      control, bool pressed, ref bool cancelDefault)` public on
      `RblClassAddIn`, exactly that three-argument signature or the menu item
      dies); the Word object model is reached by late binding through
      `AccessibleObjectFromWindow(OBJID_NATIVEOM)`, no Word interop; Ctrl+click
      in replies (own window and inline) is included; touch taps cannot be
      blocked (accepted gap, documented, no after-the-fact warning);
      banner detected from the **rendered text** at click time
      (`ExternalBannerStripper.ContainsBannerText`, tolerant to encodings),
      not from the HTML; no warning for `mailto:` or `#` anchors
      (`PhishingLinkFilter.IsWarnable`); fail-open everywhere; log only
      `scheme://host`, never the full URL; "Open anyway" re-follows through
      `Hyperlink.Follow()`. New setting `PhishingWarnLinks`, default **on**,
      in Settings next to the banner options; strings
      `PhishingLink_*` and `Settings_PhishingWarnLinks` in EN/FR/DE (table in
      study section 5.3); core types `ContainsBannerText`, `PhishingLinkFilter`,
      `PhishingWarningPolicy.ShouldWarnLink`, `LinkWarningMessage`, all
      xUnit-tested. Docs: walkthrough settings row, both `security.md` files
      (two thread-scoped hooks on Outlook's own thread, no injection, rendered
      text read at click time, no mail modified, host logged, touch gap),
      architecture hook rules. Dev verification: Claude runs the e2e cases of
      study section 5.14 (fake banner mail stored as UTF-8). **Not verifiable
      on the dev machine:** real banner/Stormshield mail, and above all the
      reaction of the corporate EDR to `SetWindowsHookEx` on the 32-bit
      target: leave the checkbox open with `- shipped, pending tests` and
      re-verify both on the real target before pilot. [#37]

---

## Phase 8 — Continuous maintenance (ongoing)

- [ ] Regression test plan executed on every Semi-Annual Enterprise channel
      update (typically January and July). Matrix: install, folder indexing,
      folder search, open-folder, classify (single + multi-destination),
      conversation widening, sub-folder creation, attachment removal, send-time
      guards, sent-item triage, settings persistence — mirrors the parity
      checklist in [docs/legacy-overview.md](docs/legacy-overview.md) §9.
- [ ] Watch the Microsoft 365 roadmap / deprecation notices for COM add-ins
      and classic Outlook.
- [ ] Annual review of the migration plan toward "new Outlook" if/when that
      becomes a corporate direction.

## Phase 9 — Security hardening (audit 2026-06-07)

A full audit found **no critical or high-severity vulnerabilities** (SQL fully
parameterized, no path traversal, no hardcoded secrets, no crypto, no outbound
network calls, COM lifetime managed via `ComRef<T>`). The items below are
**defense-in-depth and enterprise due-diligence** for a corporate security
review — several will be dropped on risk/benefit for a per-user, no-network,
single-machine add-in, but are recorded so the review trail is explicit.
Since the audit, the posture gained an opt-in update check (v2.7.0.0) and
in-box S/MIME encryption for the experimental re-encrypted filing
(v2.12.0.0); `docs/security.md` is the current reference. Completed items
are archived in [history](docs/roadmap-history.md).

### Evidence & process (highest value for sign-off)
- [ ] **Authenticode-sign the DLL + MSI** with the internal-PKI Code Signing
      cert (OID 1.3.6.1.5.5.7.3.3), timestamped — promoting today's optional
      signing to the GA default.
- [ ] **Dependency vulnerability scanning** in CI
      (`dotnet list package --vulnerable --include-transitive` + Dependabot over
      `SQLitePCLRaw`, `Microsoft.Data.Sqlite`, `Serilog`). Pin and re-verify on
      every bump — the COM host has no binding redirects
      ([[com-addin-no-binding-redirects]]).
- [ ] **SAST in CI** (Roslyn security analyzers / CodeQL, security rules as
      build errors), added to the existing `.github/workflows/ci.yml`, which
      only runs the Core tests today.
- [ ] **SBOM** per release artifact.

### Error handling & data leakage (actionable code changes)
- [ ] **Stop showing raw exception text to users.** `ShowError`
      (`RblClassAddIn.cs`) calls `MessageBox.Show(ex.ToString())`. Show a
      friendly message instead; the full exception is already logged
      (v2.7.1.0 D3).
- [ ] **Replace silent `catch { }` blocks** in `OutlookMailStore.cs` (16 left
      on 2026-11-21) with logged catches (notably the attachment-strip swallow —
      a filed copy could silently retain attachments the user believes were
      removed, [[classify-attachments-on-copy]]; surface that failure to the
      user).
- [ ] **Decide & document send-guard failure behavior.** `Application_ItemSend`
      deliberately **fails open** (CLAUDE.md forbids exceptions escaping into
      Outlook). A security team may want the external-recipient guard to **fail
      safe**. Narrow the catch so a guard *evaluation* error still prompts while
      genuine COM faults stay contained; document the stance.
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
