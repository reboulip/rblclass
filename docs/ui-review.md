# Pane UI review (v2.13.0.0 A1)

Internal study, not published on the website. It is the input for ROADMAP item
**A2** (pane redesign) and answers GitHub issue
[#67](https://github.com/reboulip/rblclass/issues/67).

- **Code studied:** branch `feature/v2.13.0.0`. The pane XAML and strings last
  changed on 2026-11-14 (commit `1657a99`); the Release build used for the
  measurements dates from 2026-11-21 and therefore matches the current code.
- **Units:** DIP = device-independent pixel (1/96 inch, WPF's unit).
  *Frame* = the Office task-pane frame (what `CustomTaskPane.Width` sets).
  *View* = the WPF `MainPaneView` inside it. *Column* = the view minus the
  root `Margin="8"` on both sides (`MainPaneView.xaml:40`), i.e. the width the
  controls actually get.
- **Status of the numbers:** every width in this document is **measured**
  (live UI Automation on the dev machine, pixel scans of the #67 screenshots,
  or the WPF layout engine run on the compiled view), not estimated from
  character counts. The few values that are extrapolated are marked
  *estimate*. Section 2.11 lists what still needs a live confirmation.

## Contents

0. [Summary](#0-summary)
1. [Inventory](#1-inventory)
2. [The #67 problem, measured](#2-the-67-problem-measured)
3. [Critique against best practices](#3-critique-against-best-practices)
4. [Font scaling versus responsive layout](#4-font-scaling-versus-responsive-layout)
5. [Layout proposals](#5-layout-proposals)
6. [Prioritized change list](#6-prioritized-change-list)
7. [Proposed A2 specification](#7-proposed-a2-specification)
8. [Decisions for the user](#8-decisions-for-the-user)
- [Appendix A: measurement method](#appendix-a-measurement-method)
- [Appendix B: natural widths](#appendix-b-natural-widths-today)
- [Sources](#sources)

---

## 0. Summary

1. **The task pane is sized in physical pixels.** `ApplyDockSize` sets
   `Width = 360` (`RblClassAddIn.cs:1221`). On the dev machine at 150 %
   scaling, the frame measures exactly 360 px = **240 DIP**, and the WPF view
   gets 344 px = **229 DIP**. That leaves a **213 DIP** column. The #67
   screenshots are 1912 px wide at 144 dpi, so the reporter's HP laptop also
   runs at **150 %**: same numbers.
2. **The starting proposal of a 260 DIP minimum cannot fix #67**, because at
   150 % today's *default* pane is already narrower than that (229 DIP view).
   **Recommended target: 200 DIP of view width** (184 DIP column, about
   316 px of frame at 150 %).
3. **Minimum view width with nothing clipped today:**

   | | EN | FR | DE |
   |---|---|---|---|
   | Header only | 223 | 221 | 223 |
   | Everything, worst case | 230 | 295 | 256 |

   At the 150 % default (229), the header fits with 6–8 DIP to spare.
   Narrowing the pane by about 10 px hides the status dot, then `?`. Already
   clipped at the default width:
   - **Undo** in FR and DE as soon as two or more destinations are checked;
   - **Options labels**: one in EN, three in FR, one in DE.
4. **The #67 screenshots show build 2.10.0.0.** Its header needed a 269 DIP
   column; v2.11.0.0 B2 cut that to 205–207 DIP. The reporter's narrow
   setting (245 DIP column) would now show the whole header, but still not
   the whole Options panel in French.
5. **The live pane renders its text in Microsoft Sans Serif 8.25 pt (11 DIP),
   not Segoe UI.** It inherits the WinForms host font through `ElementHost`'s
   default property mapping (section 1.1). This is a typography defect. It
   also means switching to Segoe UI 12 without a layout change would widen
   the worst case by about another 20 DIP (section 2.7); the header itself
   barely changes.
6. **Visible defects today:**
   - The Auto-class label breaks **mid-word** in FR and DE ("Classemen | t
     auto", "Auto- | Klassierun | g"), caused by `MaxWidth="48"` plus
     `TextWrapping="Wrap"` (`MainPaneView.xaml:96`).
   - Options labels are hard-clipped.
   - In the Sent-items triage dialog, the "Copy to Inbox (keep original…)"
     label is cut at both ends in FR and DE.
7. **Top/Bottom docking is the worst case.** The default 300 px height at
   150 % leaves a 163 DIP-tall view. The results list gets 34–49 DIP (2–3
   rows) when nothing is checked, and **0–5 DIP once destination chips
   appear**. The horizontal mode (v2.9.0.0 A2) only rearranges the action
   buttons. It does not address the real constraint, which is height.
8. **Accessibility:**
   - The pane uses no `AutomationProperties` at all.
   - UI Automation names are:
     - emoji (`💾`, `⚙️`, `?`) for the icon buttons;
     - empty for the Auto-class button, the search box and the row checkboxes;
     - the CLR type name `RBLclass.AddIn.ViewModels.SelectableFolder` for
       every result row.
   - The status dot is not exposed at all and is colour-only.
   - The header's Tab order runs right to left.
9. **WCAG contrast failures:**
   - dark-theme Classify button text: 3.01:1;
   - update banner: 2.33:1;
   - input borders against WCAG 1.4.11: 1.70:1 (light), 2.05:1 (dark);
   - status dot: yellow in the light theme, red in the dark theme;
   - the `#777` sub-headings in Settings.
10. **Recommendation:**
    - adopt **Proposal B**, a command bar with a static "⋯" overflow:
      - the index status moves into the search box;
      - redundant rows are merged;
      - every row reflows instead of clipping;
    - replace the horizontal mode with a real three-column layout;
    - use Segoe UI 12 and Segoe MDL2 Assets icons;
    - **do not** scale fonts with the pane width.

    A prototype of Proposal B's rows, built on the real styles, fits with no
    clipping and no mid-word break down to **190 DIP** (FR and DE) and
    **160 DIP** (EN).

---

## 1. Inventory

### 1.1 Host, sizing and fonts

| Item | Where | Facts |
|---|---|---|
| Task-pane creation | `RblClassAddIn.cs:1155-1194` | Created once in `EnsureTaskPane`. The dock position comes from settings (default Right, `:1167-1171`). The title is indented with 14 spaces for Top/Bottom, so the nav rail cannot hide it (`:1181-1186`). |
| Default size | `RblClassAddIn.cs:1215-1228` | Left/Right: `Width = 360`. Top/Bottom: `Height = 300`. Set only at creation. After a live redock, Office keeps its own size (`:1243-1257`, the "never resize after creation" rule). |
| Unit of `Width` | measured | Physical pixels. At 150 %: frame 360 px (240 DIP), WPF host 344 px (229 DIP). Office's frame chrome is 8 px per side horizontally, and 56 px vertically (42 px title strip plus borders). |
| WinForms host | `RblClassTaskPaneHost.cs:42` | `new ElementHost { Dock = DockStyle.Fill }`. Its `PropertyMap` is never touched. |
| Nav-rail inset | `RblClassTaskPaneHost.cs:128-162`, `NavRailInset.cs:47` | When Top/Bottom-docked, the host is left-padded by the measured rail overlap (the rail is 72 px wide at 150 %). This is 0 when docked Left/Right: in Left dock, Office reserves a 72 px strip for the rail (measured: `MsoDockLeft` = 72 + 360 px). |
| Theme | `ThemeService.cs:38-73`, `RblClassTaskPaneHost.cs:200-207` | Merges `Controls.xaml` and `Brushes.{Light,Dark}.xaml`, and binds the view's Background/Foreground. **No FontFamily or FontSize is set anywhere.** |
| Effective font | measured | Inherited text renders as **Microsoft Sans Serif at 11 DIP**, the WinForms default `Control.DefaultFont` (8.25 pt), most likely mapped onto the view by `ElementHost`'s default `Font` mapping. Evidence: the live widths of 9 controls match this font within 1.3 DIP. The `?` button is 16 DIP wide live: Microsoft Sans Serif's "?" (7.8 DIP) explains it, Segoe UI's (6.3 DIP) does not. In shot 1 of #67, the 2.10 "Classement auto" label ink is 82.0 DIP; Microsoft Sans Serif 11 predicts 81.6, Segoe UI 11 predicts 79.3. Elements with an explicit FontSize keep it (10, 11, 14), and the emoji buttons set `Segoe UI Emoji` explicitly. The modal windows are not hosted in `ElementHost`, so they most likely use Segoe UI 12 (to confirm). |

### 1.2 `MainPaneView.xaml`, row by row

Root: `Grid Margin="8" IsEnabled="{Binding IsNotBusy}"` (`:40`) with eight
rows (`:41-50`), plus a busy overlay (`:334-339`). Design-time size is
`d:DesignWidth="320"` (`:10`), wider than the real 150 % default (229).

| Row | Element | XAML | Layout and behaviour | Strings (EN) |
|---|---|---|---|---|
| 0 | Update banner | `:54-67` | Orange `#FF8C00` Border, white bold wrapping text, `Dismiss` button docked right. Shown when `HasUpdateAvailable`. | `MainPane_UpdateAvailable_Banner`, `_Dismiss` |
| 0 | Pinned-mail banner | `:69-87` | Blue `#0078D4` Border, 11 DIP indicator and italic subject (both wrap), right-aligned outlined "Delete from Inbox" button (`:81-85`, does not wrap). Shown when `HasPinnedItem`. | `MainPane_PinnedSentMail_*` |
| 0 | Header | `:88-144` | **`DockPanel LastChildFill="False"`, every child `Dock="Right"`**, declared from right to left: Auto-class, Save as, Manage attachments, Update folders, Settings, `?`, status dot. When the row is too narrow, the last-declared (leftmost) children get a zero-width slot and vanish. | see below |
| 0 | Auto-class | `:91-103` | `Padding="3,1"`. Label `TextBlock` with `FontSize="10"`, `MaxWidth="48"`, `TextWrapping="Wrap"` (`:96`). The easter-egg pig image (26×26) replaces it (`:99-100`). | `MainPane_AutoClass_Button` |
| 0 | Save as, Manage attachments, Update folders, Settings | `:104-123` | Emoji content (`💾`, `📎`, `🔄`, `⚙️`) in `Segoe UI Emoji` 14, `Padding="3,1"`, 3 DIP gaps, tooltip only. | `MainPane_*_Button`, `_ToolTip` |
| 0 | Help `?` | `:124-127` | `FontSize="14"`, inherited family (Microsoft Sans Serif), tooltip only. | `MainPane_Help_*` |
| 0 | Index status dot | `:128-143` | 10×10 `Ellipse` coloured red `#D13438`, yellow `#F2C811` or green `#3FA34D` by `IndexStatus`, tooltip `IndexStatusToolTip`. Not focusable, no automation peer. | `IndexStatus_*` |
| 1 | Search box | `:147-152` | `TextBox x:Name="QueryBox"`, tooltip only: no label, no placeholder. Handlers: Enter files, Ctrl+Enter opens, Ctrl tap toggles "list every folder", Tab opens Options (`.xaml.cs:85-116`, `:162-185`, `:240-249`). | `MainPane_QueryBox_ToolTip` |
| 2 | Results list | `:158-224` | `ListBox`, no horizontal scroll. Rows: `DockPanel` with `👁` (open, 22 wide, `:195-196`) and `⤢` (cycle path mode, 22 wide, `:197-199`) docked right, and a `CheckBox` that fills the rest (`:200-220`). Single click toggles a destination; double-click files; right-click offers Add subfolder / Rename / Delete (`:183-194`). Row height 15.5 DIP. | `MainPane_OpenFolder_ToolTip`, `_ExpandPath_ToolTip`, menu items |
| 2 | Path modes | `:203-218` | (1) one line, trimmed from the left by `PathTrim` (`RightReserve="52"`, `:204-206`; `PathTrim.cs:101-119` also keeps a 24 DIP inset); (2) wrapped (`:208-209`); (3) one segment per line (`:211-218`). Each row cycles independently (`SelectableFolder.cs:74-79`). | |
| 3 | Destinations | `:228-261` | Header row: bold `SelectionHeader` plus a "clear" button docked right (`:230-236`). Below, a `WrapPanel` of chips (rounded `Border`, name, `✕` button, full-path tooltip, `:245-257`). Only shown when at least one destination is checked. | `Status_SelectedFolders_*`, `MainPane_ClearSelection_*`, `_RemoveDestination_ToolTip` |
| 4 | Options toggle (vertical mode) / action bar (horizontal mode) | `:266-286` | Vertical: the `▶ Options` button alone (`:267-271`). Horizontal: a one-line `StackPanel` holding a second copy of Options, Classify, Undo and Journal (`:272-285`). | `Pane_Options` ("▶ Options"), `Pane_Journal` |
| 5 | Options panel | `:289-301` | Collapsed by default. Six `CheckBox`es with **non-wrapping** string content: list every folder, keep a copy, remove attachments, whole conversation, strip banner (only when a banner has been learned), mark as read. `FirstOptionCheckBox` is the Tab target. | `MainPane_*_CheckBox` |
| 6 | Classify + Undo (vertical mode) | `:303-310` | Horizontal `StackPanel`: `AccentButton` Classify (accent only when destinations exist, `Controls.xaml:396-401`) and Undo (`Padding="12,4"`). | `Status_Classify*`, `MainPane_Undo_*` |
| 7 | Feedback and journal | `:315-331` | Muted, wrapping search feedback; `▶ Journal` toggle (vertical mode only, `:318-321`); collapsible rolling journal (5 entries, `MainPaneViewModel.cs:128`). | `Status_*` |
| overlay | Busy | `:334-339` | Translucent `Theme.Background` with bold `BusyStatus`. | |

### 1.3 Layout modes

- **Vertical** (default). Rows 0–7 as above.
- **Horizontal** (v2.9.0.0 A2): `IsHorizontalLayout = Width > Height`, set on
  every `SizeChanged` (`MainPaneView.xaml.cs:32-35`,
  `MainPaneViewModel.cs:306-314`). Only rows 4, 6 and 7 change. Options,
  Classify, Undo and Journal sit on one line; row 6 and the vertical Journal
  toggle are hidden (`CanShowJournalInVerticalMode`, `:320`). Everything else
  stays stacked vertically, so the results list (the only star row) absorbs
  the whole height deficit.

### 1.4 Dock positions

| Dock | Geometry | Consequences for the pane |
|---|---|---|
| Right (default) | Frame 360 px wide, full Explorer height. | Width-bound. Narrowing the pane widens the message list (the #67 trade-off). |
| Left | Same frame width. Office places it right of the rail strip (measured on the dev machine, which is currently docked Left). | Same width analysis as Right. The pane sits between the rail and the folder pane, so mouse travel to the message list is longer. |
| Top / Bottom | Full Explorer width minus the rail inset; frame 300 px tall at creation, i.e. a **163 DIP** view at 150 %. | Height-bound; the horizontal mode applies. The header is pushed to the far right edge, ~1000 DIP away from the action bar at the left. Measured in 2.8. |
| Floating | Any size. | `Width > Height` decides the mode; both modes must survive any aspect ratio. |

### 1.5 Modals

All modals are owned by the Outlook window and centred on it
(`DialogPlacement.cs:19-34`). Each calls `ThemeService.Apply` from its
code-behind.

| Window | Size, resize | Content | Notes |
|---|---|---|---|
| Settings | 700×560, resizable, min 580×540 (`SettingsWindow.xaml:8-10`) | Left nav `TabControl` (8 tabs, nav column min 130, `Controls.xaml:509`); grouped settings with ⓘ tooltips; Close. | Checkbox labels already wrap through `TextBlock` content with a `MaxWidth`: the pattern the pane should reuse. Two sub-headings hard-code `Foreground="#777"` (`:100`, `:112`). |
| Attachment disposition | 760×540, resizable, min 560×340 (`AttachmentDispositionWindow.xaml:9-10`) | Intro, notices, three mode radios, shared folder picker, strip toggle, grouped panel or per-attachment rows (Delete / Save to / Keep, per-row folder picker); Confirm / Cancel. | Fine at its size. |
| Save-as picker | 560×320, resizable, min 400×220 (`SaveAsPickerWindow.xaml:9-10`) | Intro, favourite-folder search with Browse, result list (max 140 tall), target link; Save / Cancel. | Fine. |
| Sent-item triage | 380×330, **not resizable** (`SentItemTriageWindow.xaml:8-9`) | Bold question, muted subject, five full-width buttons. | Button text does not wrap. "Copy to Inbox (keep original…)" needs 399 DIP in FR and 351 in DE against about 322 available (the window frame width is assumed), and the centred content is cut at **both** ends. |
| Name prompt | 360×160, not resizable (`NamePromptWindow.xaml:8-10`) | Label, text box, Create / Cancel. | Fine for current strings. |
| Meeting hint | 420×230, not resizable (`MeetingClassifyHintWindow.xaml:8-10`) | Bold line, body, "don't show again", OK. | Fits by *estimate* (about 154 of about 191 DIP of client height in FR/DE); no `SizeToContent` safety. |
| About (ribbon) | 400×280, not resizable | Version, author, documentation link. | Out of scope; listed for completeness. |

---

## 2. The #67 problem, measured

### 2.1 The report

Issue [#67](https://github.com/reboulip/rblclass/issues/67), in French:

> "Avec la nouvelle gamme de PC HP, l'écran est légèrement plus petit que sur
> les Lenovo […] la répartition des colonnes entre Outlook et RBL class peut
> masquer des options RBL ou alors ne pas permettre l'affichage colonne des
> mails."

The reporter suggests a font size proportional to the pane. The owner added
that a more efficient and ergonomic design should be considered as well.

What the two screenshots show (pixel scans, 150 %):

| | Shot 1 (wide pane) | Shot 2 (narrow pane) |
|---|---|---|
| Pane frame | ≈ 467 px = 311 DIP | ≈ 406 px = 271 DIP |
| Search box = content column | 429 px = **286 DIP** | 368 px = **245 DIP** |
| Header | complete: dot, gear, refresh, clip, save, "Classement auto" = 404 px = 269 DIP | dot hidden, gear cut by 8 DIP: needs 245 + 24 = 269 DIP |
| Message list | ≈ 1030–1070 px (687–714 DIP): **compact 2-line layout** | ≈ 1093–1130 px (729–755 DIP): **column view** |

The screenshots predate v2.11.0.0 B2. There is no `?` button, and the
2.10.0.0 header used `Padding="6,1"`, 6 DIP gaps and an 11 DIP one-line
label (`git show 2.10.0.0:…/MainPaneView.xaml`). Today's header is 62 DIP
narrower. The underlying layout problem has not changed, though: right-docked
items vanish without any indication.

### 2.2 Default width per display scaling

`Width = 360` is applied in physical pixels, so the DIP width of the default
pane shrinks as scaling grows. The 16 px chrome was measured at 150 % only;
the other rows assume the same value.

| Windows scaling | Frame (DIP) | View (DIP) | Column (DIP) | Today's header (needs 205–207 DIP column) |
|---|---|---|---|---|
| 100 % | 360 | 344 | 328 | fits |
| 125 % | 288 | 275 | 259 | fits |
| **150 % (dev machine, HP target)** | **240** | **229** | **213** | fits, 6–8 DIP spare |
| 175 % | 206 | 197 | 181 | dot hidden, `?` cut, at the *default* width |
| 200 % | 180 | 172 | 156 | dot, `?` hidden, gear cut |

A 14″ FHD laptop at 150 % offers a 1280 DIP-wide desktop. A 15.6″ FHD laptop
at 125 % offers 1536 DIP. That 17 % loss is the most likely reason the HP
fleet feels "slightly smaller" (*hypothesis*: the Lenovo scaling is unknown).

### 2.3 Method and error margin

Three independent sources (details in [Appendix A](#appendix-a-measurement-method)):

1. **Live pane** on the dev machine (Outlook running, minimized, docked Left,
   FR UI, easter egg active). Read with Win32 window rectangles and UI
   Automation from a separate process; nothing was clicked, focused, resized
   or redocked.
2. **#67 screenshots**: pixel scans.
3. **Offline layout harness**: the compiled Release `MainPaneView` loaded in
   a separate PowerShell process (never touching Outlook), with the real
   `Controls.xaml` and brush dictionaries, the real `LocalizationService` in
   EN/FR/DE, and a stub DataContext. The WPF layout engine measures each
   control. With the live font reproduced (Microsoft Sans Serif 11 on the
   view), it matches the live pane:

| Element (FR) | Live (DIP) | Harness (DIP) |
|---|---|---|
| Save as button | 27.3 | 27.2 |
| `?` button | 16.0 | 15.8 |
| Classify ("Classer") | 62.0 | 63.3 |
| Undo ("Annuler") | 63.3 | 64.0 |
| `▶ Options` / `▶ Journal` | 60.7 / 58.7 | 60.4 / 58.5 |
| Clear ("effacer") | 46.7 | 47.6 |
| Row buttons | 22.0 × 14.0 | 22.0 × 14.5 |
| Result row height | 15.3 | 15.5 |
| Content column | 213.3 | 213.3 |

**Margin of error:** ±1.5 DIP per control (worst observed 1.3), about
±3 DIP on a full header row. Theme does not affect layout: both brush
dictionaries only define colours.

### 2.4 Header arithmetic (today)

Each value is the natural width in DIP, including the control's right margin.

| Element (XAML) | Composition | EN | FR | DE |
|---|---|---|---|---|
| Status dot (`:128`) | 10 + 3 margin | 13.0 | 13.0 | 13.0 |
| `?` (`:124`) | glyph 7.8 + padding 6 + border 2 + margin 3 | 18.8 | 18.8 | 18.8 |
| `⚙️` `🔄` `📎` `💾` (`:104-123`) | (glyph 19.2 + 6 + 2 + 3) × 4 | 120.8 | 120.8 | 120.8 |
| Auto-class (`:91-103`) | label ≤ 48 at 10 DIP + 8 | 54.7 (1 line) | 52.5 (2 lines, mid-word) | 54.8 (3 lines, mid-word) |
| **Minimum column** | sum | **207.3** | **205.1** | **207.4** |
| **Minimum view** | + 2 × 8 | **223** | **221** | **223** |

The widest single element is the Auto-class text button, and it is also
broken. "Classement" (52.9 DIP) and "Klassierung" (52.4 DIP) are wider than
the 48 DIP box, so `TextWrapping="Wrap"` splits them mid-word. The DE header
grows to three text lines (38 DIP tall).

### 2.5 Minimum view width per state (today)

| State | EN | FR | DE | What clips first below that width |
|---|---|---|---|---|
| Nothing checked, Options closed | 223 | 221 | 223 | status dot, then `?` |
| 2 destinations checked | 223 | 233 | 244 | FR/DE: Undo, after "Classer dans 2 dossiers" / "In 2 Ordner klassieren" |
| 12 destinations checked | 223 | 240 | 251 | Undo |
| Options open (and 2 destinations) | 230 | **295** | 255 | longest Options label (see below) |
| Update and pinned banners, journal | 223 | 221 | 223 | header (the FR pinned "Supprimer de la boîte de réception" needs 214) |
| Horizontal bar (12 destinations, journal) | 337 | 371 | 388 | Journal button at the end of the one-line bar |
| **Worst case, vertical** | **230** | **295** | **256** | Options label |

The longest Options labels:

- EN "Remove attachments from the filed copy": 214.3 DIP column;
- FR "Conserver une copie (ne pas supprimer les originaux)": 279.0 DIP column;
- DE "Anhänge aus der abgelegten Kopie entfernen": 239.5 DIP column.

### 2.6 What breaks at today's 150 % default (view 229, column 213)

- **Header:** fits, with 6.0 (EN), 8.2 (FR) and 5.9 (DE) DIP spare. Any
  narrowing beyond ~9–12 px hides the dot, then `?`.
- **Undo:** clipped in FR, by 4 DIP with 2 destinations and 11 DIP with 12.
  Clipped in DE by 15–21 DIP.
- **Options labels, hard-clipped:**
  - EN: "Remove attachments from the filed copy" (1 DIP);
  - FR: "Supprimer la bannière d'expéditeur externe" (228), "Supprimer les
    pièces jointes de la copie classée" (252), and "Conserver une copie (ne
    pas supprimer les originaux)" (279). In the last one the cut lands on the
    **negation**: "(ne pas supp…";
  - DE: "Anhänge aus der abgelegten Kopie entfernen" (240).
- **Auto-class:** the mid-word break, in FR and DE, at every width.

### 2.7 Switching the font alone would make things worse

The same harness with Segoe UI 12 (what Outlook uses) and no layout change
gives:

| | EN | FR | DE |
|---|---|---|---|
| Header | 221 | 222 | 219 |
| Worst case | 249 | 315 | 275 |

The font fix therefore has to ship with the layout fix, never on its own.

### 2.8 Top/Bottom docking: the height budget

The view is about 1200 DIP wide, so `IsHorizontalLayout` is true. The table
shows the results-list height in DIP for EN / FR / DE; one row is 15.5 DIP.

| State | Height 163 (150 %, default 300 px) | 202 (125 %, *estimate*) | 263 (100 %, *estimate*) |
|---|---|---|---|
| Nothing checked, Options closed | 49 / 45 / 34 (2–3 rows) | 88 / 84 / 73 | 149 / 145 / 134 |
| 2 destinations (chips shown) | **5 / 1 / 0** | 44 / 40 / 29 | 105 / 101 / 90 |
| Options open and chips | **0**, 5–6 checkboxes pushed below the edge | 0, 3–4 cut | 0–2, none cut |

At 150 % the chips alone wipe out the results list: the user can no longer
see the folders they are choosing from. The one-line action bar saves about
30 DIP, while the stacked header, search box, chips and Options panel cost
120–250 DIP. The horizontal mode must become a multi-column layout (section
5.4).

### 2.9 The Outlook side: the compact-layout threshold

Outlook switches the message list to its compact two-line layout when the
list is narrower than **View > View Settings > Other Settings > "Use compact
layout in widths smaller than N characters"**. 125 characters is the commonly
reported value; that is a forum value, not an official default
([Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/4663823/how-to-restore-default-outlook-view),
[ExtendOffice](https://www.extendoffice.com/outlook-receive-read-emails/outlook-inbox/change-inbox-view/other-settings.html)).
The screenshots place the switch between about 690 and 730 DIP, which is
consistent with that setting.

Two consequences:

- **Mitigation available today, no code:** lowering N on the HP machines
  (for example to 100) keeps the column view with a narrower message list.
  Hiding or bottom-docking the reading pane and narrowing the folder pane
  also help.
- **Budget:** 1280 DIP − 48 rail − about 233 folder pane − about 700 list −
  about 10 splitters leaves about 290 DIP for the pane frame. The current
  default (240) already fits that budget; the pane must not get *wider* by
  default at 150 %.

### 2.10 Target minimum width

**Recommendation: 200 DIP of view width (184 DIP column).** The frame width
this needs is 200 × scale + 16 px:

| Scaling | 100 % | 125 % | 150 % | 175 % | 200 % |
|---|---|---|---|---|---|
| Frame at the 200 DIP target | 216 px | 266 px | **316 px** | 366 px | 416 px |

Why 200:

- At 150 % it leaves 44 px (12 %) of headroom below the default, against
  ~10 px today.
- With Proposal B, the narrowest no-clip width is 190 DIP in FR/DE, so the
  target keeps a 10 DIP safety margin and still leaves about 15 characters of
  folder path per row.
- 180 DIP would cost legibility: about 11 characters of path, and two-line
  buttons in FR/DE.
- **260 DIP is rejected:** it is wider than today's default view at 150 %,
  so it would force users back to the wide pane that triggers the compact
  message list.

At 175 % and 200 %, the 360 px default would fall *below* the target. A
creation-time floor fixes that without touching 150 %:
`Width = max(360, ceil(229 × scale) + 16)`, which gives 360 px at 100–150 %,
417 px at 175 % and 474 px at 200 %. This change is applied in
`ApplyDockSize` only, before `Visible` (C16).

### 2.11 To confirm live

1. `CustomTaskPane.Width` / `Height` units and the chrome size at 100 % and
   125 %, and on the 32-bit target. The dev machine (150 %, ARM64 under
   emulation) shows physical pixels. Log the requested size, the DPI, and the
   host's first pixel size (C16).
2. The pane font: log `FontFamily` and `FontSize` of the view after the
   `ElementHost` attaches it, or inspect with a visual-tree tool. The evidence
   points to Microsoft Sans Serif 8.25 pt.
3. The compact-layout N on the HP machines, and their reading-pane and
   folder-pane widths.
4. The Top/Bottom chrome at 100 % and 125 % (the 202 and 263 DIP columns
   above are estimates).
5. A2's final layout at 200 DIP and at the default, in EN/FR/DE,
   light and dark, Left/Right and Top/Bottom. This is required by A2's
   acceptance criteria.

---

## 3. Critique against best practices

Severity: **H** high (breaks function or access), **M** medium, **L** low.

### 3.1 Visual hierarchy and primary action

| # | Element | Guideline | Finding | Impact | Sev |
|---|---|---|---|---|---|
| F1 | Auto-class (`:91-103`) | Fluent hierarchy: the most important actions get the most prominent treatment; one type ramp | A core filing action has the *smallest* text in the pane (10 DIP, squeezed into 48 DIP, split over 2–3 lines and mid-word in FR/DE), next to six same-weight icons. | Reads as an afterthought; looks broken in FR/DE. | M |
| F2 | Search box (`:147-152`) | Label every input; a placeholder signposts the entry point | No label, no placeholder, only a hover tooltip. The pane's primary input looks like an empty bar. | First-use discoverability; screen readers announce an unnamed edit field. | M |
| F3 | Classify (`:303-310`, `Controls.xaml:361-435`) | One prominent primary action | Correctly the only accent button, and accent only once destinations exist (good). It sits below the Options panel, so opening Options moves it down by ~100 DIP. | Minor target shift. | L |
| F4 | Destinations header (`:230-236`) and Classify label (`MainPaneViewModel.cs:215-221`) | Remove redundancy | "Selected: 2 folders" and "Classify to 2 folders" say the same thing. The header costs a row and is the widest text after the Options labels (FR 133 DIP + 48 DIP button). | Wasted height and width. | L |

### 3.2 Progressive disclosure

| # | Element | Guideline | Finding | Impact | Sev |
|---|---|---|---|---|---|
| F5 | Options and Journal toggles (`:267-271`, `:318-321`) | Disclosure controls expose their state (Fluent Expander or ToggleButton; WCAG 4.1.2) | Plain Buttons with "▶" baked into the string (`Pane_Options`). The glyph never turns to "▼", and UI Automation sees no expanded or toggled state. | State invisible, especially to screen readers. | M |
| F6 | Header | Fluent command bar: primary commands visible, secondary ones in a "More" overflow | Rare commands (Update folders, Settings, Help) permanently take 79 DIP of a row that the frequent commands need at narrow widths. | The root of the width budget problem. | H |
| F7 | Status dot | Show status when it matters, with text | A permanent dot even when all is fine; the explanation and the remedy ("click Refresh folders") only appear on hover. | Red/yellow states go unexplained. | M |

### 3.3 Icon-only buttons

| # | Element | Guideline | Finding | Impact | Sev |
|---|---|---|---|---|---|
| F8 | All icon buttons, search box, result rows | WCAG 4.1.2 Name, Role, Value; Fluent: icon-only buttons need a tooltip **and** an accessible name | No `AutomationProperties` anywhere in the add-in (grep). Live UI Automation names: `💾` `📎` `🔄` `⚙️` `?` `👁` `⤢` `✕`; **empty** for the Auto-class button (its content is a Grid), the search box and the row checkboxes; **`RBLclass.AddIn.ViewModels.SelectableFolder`** for every result row (the `ToString` fallback). HelpText falls back to the tooltip, which is good. | Screen readers say "floppy disk", "button", or read out a type name. | H |
| F9 | Status dot (`:128-143`) | WCAG 1.4.1 Use of Color, 4.1.2 | An `Ellipse` has no automation peer, so it is absent from UI Automation. It is not focusable, and its meaning rests on colour only. | Status unreachable by keyboard and screen readers. | H |
| F10 | Emoji glyphs | Fluent iconography: one system icon font (Segoe MDL2 Assets on Windows 10, Segoe Fluent Icons on 11), consistent size and stroke | WPF on .NET Framework 4.8 renders emoji as monochrome outlines (`:14-16`). Weights are uneven; `⚙️` carries a variation selector; `👁` and `⤢` come from different fallback fonts (live ink: ⤢ 8 DIP, 👁 14 DIP). | Visual noise, inconsistent sizes, rendering depends on font fallback. | M |
| F11 | `?` (`:124-127`) | Consistency | Inherits Microsoft Sans Serif, unlike its emoji neighbours. | Minor mismatch. | L |

### 3.4 Click targets

| # | Element | Guideline | Finding | Impact | Sev |
|---|---|---|---|---|---|
| F12 | Row buttons, rows, chip `✕` | WCAG 2.2 SC 2.5.8 Target Size (Minimum): 24×24 CSS px (AA); Windows touch guidance about 40 epx | Row buttons 22 × 14.5 DIP, 3 DIP apart (fails, and the spacing exception does not apply). Rows 15.5 DIP tall. Chip `✕` about 13 × 12.5 DIP. | Easy to hit "expand" instead of "open"; hard on touch laptops. | M |
| F13 | Header buttons | Comfortable pointer targets (≥ 32 epx recommended) | 27 × 23 DIP; `?` 15.8 DIP wide. | Acceptable but small. | L |

### 3.5 Spacing, alignment and typography

| # | Element | Guideline | Finding | Impact | Sev |
|---|---|---|---|---|---|
| F14 | Whole pane | Fluent 4 epx spacing grid | Spacing values of 2, 3, 4, 6 and 8 DIP are mixed (3 DIP header gaps `:104-128`, 6 DIP row gaps `:52`, `:147`, `:228`). | Irregular rhythm. | L |
| F15 | Whole pane | One family, one type ramp; match the host application | Inherited Microsoft Sans Serif 11 (section 1.1), plus explicit 10, 11 and 14, plus Segoe UI Emoji. Outlook's UI uses Segoe UI at about 9 pt (12 DIP). | The pane looks foreign to Outlook; 10 DIP text is small on a 14″ screen at 150 %. | M |
| F16 | `d:DesignWidth="320"` (`:10`) | Design at the real size | The real default at 150 % is 229. | The designer preview hides every problem in this document. | L |

### 3.6 Text truncation

| # | Element | Guideline | Finding | Impact | Sev |
|---|---|---|---|---|---|
| F17 | Options labels (`:291-300`, CheckBox template `Controls.xaml:55-56`) | Fluent: wrap long labels; ellipsize only with a tooltip; never hard-clip | Hard-clipped mid-word, with no ellipsis and no wrap. | Meaning lost: "(ne pas supp…" hides the negation of a destructive option. | H |
| F18 | Auto-class label (`:96`) | No mid-word breaks | `Wrap` with `MaxWidth="48"`. `WrapWithOverflow` (or more room) prevents it. | Looks broken. | M |
| F19 | Classify + Undo (`:303-310`), horizontal bar (`:272-285`) | Reflow before clipping | Horizontal `StackPanel`s clip the right-hand buttons. | Undo half-visible in FR/DE. | H |
| F20 | Path trimming (`PathTrim.cs`) | Ellipsize with a tooltip | Leading-ellipsis trimming with a full-path tooltip (good). The reserve (52 + 24 DIP) exceeds the 47 DIP button block plus scrollbar. | About 12 DIP of path lost. | L |
| F21 | Chips (`:245-257`) | Bound long items | No `MaxWidth` or trimming: a long folder name makes a chip wider than the pane. | Clipped chip, unreachable `✕`. | M |
| F22 | Header `DockPanel` (`:88`) | Never hide commands without an affordance | The right-docked chain silently gives the leftmost items a zero-width slot. | The #67 symptom. | H |

### 3.7 Keyboard flow

| # | Element | Guideline | Finding | Impact | Sev |
|---|---|---|---|---|---|
| F23 | Shortcuts and focus return | Keep what works | `/` through `PreviewTextInput` (`.xaml.cs:221-227`), `ReturnFocusToSearch` after every action (`:130-135`), Ctrl+Z (`:194-213`), Enter / Ctrl+Enter, Ctrl tap: solid, and documented in the walkthrough. | Must survive A2 unchanged. | – |
| F24 | Header | WCAG 2.4.3 Focus Order | Tab runs right to left (Auto-class → … → `?`), because the `DockPanel` children are declared in reverse visual order. | Disorienting. | M |
| F25 | Search box → results | Predictable focus path; arrow keys in search-as-you-type lists | Tab from the search box opens Options and jumps to its first checkbox (`.xaml.cs:173-184`), so Tab never reaches the results list. There is no Up/Down from the box to move the highlight. The list is only reachable with Shift+Tab from the controls below it. | Choosing a result other than the first needs the mouse. | M |
| F26 | Focus ring (v2.11.0.0 B4) | WCAG 2.4.7 Focus Visible | Implemented for Button, CheckBox, RadioButton, TextBox, ComboBox, ListBoxItem and TabItem (`Controls.xaml:29-122`, `136-219`, `224-297`, `313-435`, `455-495`). **No style** exists for ToggleButton, ContextMenu/MenuItem or Expander; the row context menu (`:183-194`) uses the default menu chrome (dark mode to be checked). | Any redesign that adds these controls must add themed templates with the same ring. | M |

### 3.8 Contrast (WCAG 2.x: 4.5:1 for text, 3:1 for UI component boundaries)

| # | Pair | Light | Dark | Verdict |
|---|---|---|---|---|
| – | Body text on background | 16.46 | 12.66 | pass |
| – | Muted text (feedback, journal) | 6.46 | 7.14 | pass |
| – | Text on button and selection backgrounds | 14.72 / 12.54 | 10.02 / 6.59 | pass |
| – | Focus ring vs background / button | 4.53 / 4.05 | 6.14 / 4.86 | pass (3:1) |
| – | Pinned banner, white on `#0078D4` | 4.53 | 4.53 | pass, barely |
| – | Classify text (accent) | 4.53 (white on `#0078D4`) | – | pass, barely |
| F27 | Classify text, white on `#2899F5` (dark accent) | – | **3.01** | **fail** 1.4.3. Fix: `#115EA3` gives 6.66 |
| F28 | Update banner, white bold 11 DIP on `#FF8C00` | **2.33** | **2.33** | **fail**. Fix: text `#201F1E` gives 7.05 |
| F29 | Text box and list borders (`#C8C6C4`, `#5A5A5A`) | **1.70** | **2.05** (1.62 against its own fill) | **fail** 1.4.11. Fix: `#8A8886` light (3.53), `#A19F9D` dark (4.25 on its fill) |
| F30 | Status dot: yellow (light), red (dark) | **1.61** | **2.87** | **fail** 1.4.11, and colour-only (F9) |
| F31 | Settings `#777` italic sub-headings (`SettingsWindow.xaml:100`, `:112`) | **4.48** | **3.16** | **fail**. Fix: `MutedText` style |

### 3.9 Modals

| # | Element | Finding | Sev |
|---|---|---|---|
| F32 | Sent-items triage (`SentItemTriageWindow.xaml:24-34`) | Fixed 380 DIP width with non-wrapping button text. FR overflows by 77 DIP and DE by 29, and the centred label is cut at both ends. Fix: wrapping `TextBlock` content plus `SizeToContent="Height"`. | H |
| F33 | Fixed-size dialogs (name prompt, meeting hint, triage) | No `SizeToContent` safety. Current strings fit except F32, but longer translations would clip vertically. | L |
| F34 | Settings | Good wrapping and ⓘ patterns; the only issue is F31. | – |

---

## 4. Font scaling versus responsive layout

The issue suggests a font size proportional to the pane width. This was
assessed against three responsive techniques: wrapping, an overflow menu and
icon-only modes.

**How it would be built in WPF 4.8:** wrap the view in a `Viewbox`
(`Stretch="Uniform"`, `StretchDirection="DownOnly"`), or bind `FontSize` to
`ActualWidth` through a converter.

**Why not:**

1. **It shrinks text exactly where the screen is already smallest.** At
   150 % on a 14″ FHD panel, an 11 DIP font has an em size of about 2.7 mm.
   Narrowing the pane by 20 % would bring it to about 9 DIP (6.6 pt). That is below
   Fluent's smallest ramp size (12 epx caption) and below Office's 9 pt UI
   text. WCAG 1.4.4 (Resize Text) and the Windows "Text size" setting assume
   text that the user, not the window width, controls.
2. **`Viewbox` scales everything, not just text.** Hit targets shrink with
   it, which worsens F12. One-pixel borders and the focus ring go blurry or
   sub-pixel, and pixel snapping is lost. The vertical axis scales too:
   `Uniform` must keep the aspect ratio, so the results list gets *shorter*
   when the pane gets narrower.
3. **Popups do not follow.** Tooltips, context menus and modals live outside
   the visual tree and keep their own size, so the pane would look
   inconsistent with its own menus.
4. **It does not remove the cause.** Seven fixed items in a right-docked
   chain still overflow; scaling only moves the breaking point and makes the
   whole UI smaller to save room for rarely used buttons.
5. **Every width becomes a different design.** A2's live check at one
   target width would no longer cover the other widths.

**Responsive techniques available in .NET Framework 4.8, with no new
dependency:**

- `Grid` star columns with `MinWidth` (elastic regions);
- `WrapPanel` (rows reflow onto a second line instead of clipping);
- `TextWrapping="Wrap"` or `"WrapWithOverflow"` (labels grow taller, never
  cut mid-word);
- a static "⋯" overflow (`Button` plus `ContextMenu`) for rare commands;
- icon-only commands with tooltip and accessible name, only for universally
  understood glyphs (save, attach, more, undo);
- `DataTrigger`s on a view-state property, to re-place the same elements in
  another arrangement (the horizontal mode).

**Recommendation:** keep **one fixed font size** (Segoe UI 12 DIP, 11 for
secondary text). Make the layout responsive with the techniques above. Reject
proportional font scaling and `Viewbox`. If users want larger text, that is
the job of Windows' "Text size" or scaling settings, which WPF already honours.

---

## 5. Layout proposals

### 5.1 Shared building blocks (both proposals)

- **No right-docked `DockPanel` chains.** Rows use `Grid` (`*` plus `Auto`)
  or `WrapPanel`. Children are declared in visual order, so the Tab order is
  left to right.
- **Wrap, never clip.** Options labels, the pinned-banner button, the
  Classify label and the triage buttons get `TextBlock` content with
  `TextWrapping="Wrap"`. The Auto-class label gets `WrapWithOverflow` and
  loses its `MaxWidth`.
- **One instance of each control.** Today Options, Classify, Undo and Journal
  exist twice (rows 4 and 6). The horizontal mode should re-place the same
  elements through `Style` setters on `Grid.Row` / `Grid.Column`, driven by
  `IsHorizontalLayout`, so `x:Name`s, focus and event handlers stay unique.
- **Accessible names** on every icon-only control, the search box (name and
  placeholder) and the result rows (`AutomationProperties.Name="{Binding
  DisplayPath}"`). Live regions (`AutomationProperties.LiveSetting="Polite"`,
  available since .NET Framework 4.7.1) on the search feedback and status
  notice. WPF does not raise the change event by itself: the code must call
  `RaiseAutomationEvent(AutomationEvents.LiveRegionChanged)` on the
  element's peer when the text changes.
- **Contrast fixes** F27–F31.
- **Top/Bottom:** a three-column arrangement (5.4).
- **Rejected alternative:** a root `MinWidth` with a horizontal
  `ScrollViewer` below it. It technically avoids clipping, but parts of the
  header can then be scrolled out of view, contrary to A2's "never pushed out
  of view" criterion. It also turns a narrow pane into a scrolling chore.

Mockup legend: `[S]` save as, `[A]` manage attachments, `[...]` more menu,
`[<-]` undo, `[R]` update folders, `[G]` settings, `[?]` help, `o` / `(o)`
index status dot, `[ex]` path-display toggle, `[op]` open folder, `x` remove
chip, `[X]` clear all destinations. Mockups are schematic, at about 6 DIP per
character.

#### Reference: today, FR, 229 DIP view (150 % default)

```
+------------------------------------+
|o [?][G][R][A][S] [Classemen]       |  <- label split mid-word
|                  [t auto   ]       |
|[__________________________________]|  <- no label, no placeholder
| ... results ...                    |
|Sélectionnés : 2 dossiers  [effacer]|
|(Client A x) (Contrats 2026 x)      |
|[> Options]                         |
|[ ] Conserver une copie (ne pas sup |  <- hard clip, negation lost
|[ ] Supprimer les pièces jointes de |  <- hard clip
|[Classer dans 2 dossiers] [Annul    |  <- Undo clipped (FR, 2+ destinations)
|3 dossiers.                         |
|[> Journal]                         |
+------------------------------------+
```

### 5.2 Proposal A: Reflow (everything stays visible)

Minimal structural change: same rows, same commands, same order. Every row
flows instead of clipping.

**At the target width, 200 DIP view, FR, 2 destinations, Options open:**

```
+--------------------------------+
|o [?][G][R][A][S]               |
|[Classement auto]               |  <- WrapPanel: wraps to a 2nd line
|[______________________________]|
|+------------------------------+|
||[ ] ...s / Client A   [ex][op]||
||[v] ...Contrats 2026  [ex][op]||
||                              ||
|+------------------------------+|
|Sélectionnés : 2       [effacer]|  <- Grid: header text wraps
|dossiers                        |
|(Client A x) (Contrats 2026 x)  |
|[> Options]                     |
|[ ] Conserver une copie (ne pas |  <- label wraps
|    supprimer les originaux)    |
|[Classer dans 2 dossiers]       |
|[Annuler]                       |  <- WrapPanel: Undo wraps
|3 dossiers.                     |
|[> Journal]                     |
+--------------------------------+
```

**At a wide width, 360 DIP view (100 % default), FR:**

```
+----------------------------------------------------------+
|o [?][G][R][A][S] [Classement auto]                       |
|[________________________________________________________]|
|+--------------------------------------------------------+|
||[ ] Archive / Projets / Client A                [ex][op]||
||[v] ... / Projets / Client B / Contrats 2026    [ex][op]||
||                                                        ||
|+--------------------------------------------------------+|
|Sélectionnés : 2 dossiers                        [effacer]|
|(Client A x) (Contrats 2026 x)                            |
|[> Options]                                               |
|[ ] Conserver une copie (ne pas supprimer les originaux)  |
|[Classer dans 2 dossiers] [Annuler]                       |
|3 dossiers.                                               |
|[> Journal]                                               |
+----------------------------------------------------------+
```

- **Moves:** nothing between rows. The header becomes a left-to-right
  `WrapPanel` in today's visual order. Undo wraps under Classify when needed.
- **Merged:** nothing.
- **Dropped:** the right-docked header `DockPanel`; the 48 DIP `MaxWidth` on
  the Auto-class label.
- **Cost:** with Segoe UI 12 and 28 DIP icon buttons, the header only fits
  on **one** line from a 261 (EN), 294 (FR) or 297 (DE) DIP view. So it takes
  **two lines at today's 150 % default and at the target**. With today's
  27 DIP emoji buttons and 10 DIP label it stays on one line at 229, but keeps
  F1, F10 and F13.

### 5.3 Proposal B: Command bar with overflow (recommended)

Frequent mail actions stay visible, rare commands move behind "⋯", the status
moves to where it matters, and redundant rows merge.

**At the target width, 200 DIP view, FR, 2 destinations, Options open:**

```
+--------------------------------+
|[Classement]      [S] [A] [...] |  <- label wraps at a word boundary
|[auto      ]                    |
|[Rechercher un dossier...   (o)]|  <- placeholder, status inside
|+------------------------------+|
||[ ] ...s / Client A   [ex][op]||
||[v] ...Contrats 2026  [ex][op]||
||                              ||
|+------------------------------+|
|(Client A x)                [X] |  <- chips + clear-all icon
|(Contrats 2026 x)               |
|v Options             > Journal |  <- two ToggleButtons, one row
|[ ] Conserver une copie (ne pas |
|    supprimer les originaux)    |
|[ Classer dans 2 dossiers ][<-] |  <- Classify stretches, Undo icon
|3 dossiers.                     |
+--------------------------------+
```

**At a wide width, 360 DIP view (100 % default), FR:**

```
+----------------------------------------------------------+
|[Classement auto]                           [S] [A] [...] |
|[Rechercher un dossier...                             (o)]|
|+--------------------------------------------------------+|
||[ ] Archive / Projets / Client A                [ex][op]||
||[v] ... / Projets / Client B / Contrats 2026    [ex][op]||
||                                                        ||
|+--------------------------------------------------------+|
|(Client A x) (Contrats 2026 x)                         [X]|
|> Options                                       > Journal |
|[              Classer dans 2 dossiers              ][<-] |
|3 dossiers.                                               |
+----------------------------------------------------------+
```

**The "⋯" menu:** Update folders · Settings… · Online documentation · ─ ·
"● Folder index ready" (status line, not clickable). When the index is not
ready, a muted notice line appears under the search box with the
`IndexStatusToolTip` text, as a polite live region.

- **Moves:**
  - Update folders, Settings and Help go into the "⋯" menu;
  - the status dot goes inside the search box, with a notice line and a menu
    line;
  - Journal moves to the Options row (right side);
  - Undo becomes an icon next to Classify;
  - "clear" becomes an icon next to the chips.
- **Merged:**
  - "Selected: N folders" into the Classify label; `SelectionHeader` is kept
    as the chips' accessible name;
  - the Options and Journal toggles share a row;
  - the duplicated vertical and horizontal copies of Options, Classify, Undo
    and Journal become single instances.
- **Dropped:**
  - the header `DockPanel`;
  - "▶" baked into strings (replaced by a chevron glyph with real toggle
    state);
  - the 10 DIP squeezed Auto-class label;
  - emoji icons (replaced by Segoe MDL2 Assets);
  - the one-line horizontal action bar (5.4).

**Prototype measurement.** The rows were built as loose XAML on the real
`Controls.xaml` styles, with Segoe UI 12 and MDL2 icons (Appendix A):

| | EN | FR | DE |
|---|---|---|---|
| Narrowest view with no clipping and no mid-word break | 160 | 190 | 190 |
| Problems at 200 and at 229 | 0 | 0 | 0 |

What binds at the narrowest width is the longest word of the Auto-class label
("Classement" 60.4 DIP, "Klassierung" 60.2 DIP) plus the three 28 DIP icons.
Row heights at 200 DIP in FR: command bar 42 (two-line label; 28 in EN),
search 30, chips 80, toggles 24, Classify 28 (it still fits on one line:
"Classer dans 12 dossiers" is 129 of 130 DIP). The chips row is three short
lines only because the prototype used a text "Tout effacer" button; the
clear-all icon of the final design saves about one line. The Auto-class
label returns to one line from a 218 DIP view in FR and 221 in DE, so it is on
one line at today's 150 % default.

### 5.4 Top/Bottom docking (both proposals): three columns

Replace v2.9.0.0 A2's one-line bar. Same elements, re-placed when
`IsHorizontalLayout`. New rule: `ActualWidth >= 560 && ActualWidth >
ActualHeight`; below 560 the vertical layout is kept.

```
+--------------------------------+--------------------------------+----------------------------------+
|[Classement auto]   [S][A][...] | v Options                      |(Client A x) (Contrats 2026 x) [X]|
|[Rechercher un dossier...   (o)]| [ ] Lister tous les dossiers   |+--------------------------------+|
|[ Classer dans 2 dossiers ] [<-]|     correspondants             ||[ ] .../Client A        [ex][op]||
|3 dossiers.          > Journal  | [ ] Conserver une copie (ne pas||[v] .../Contrats 2026   [ex][op]||
|                                |     supprimer les originaux)   ||[ ] .../Archive 2025    [ex][op]||
|                                |     (scrolls when needed)      |+--------------------------------+|
+--------------------------------+--------------------------------+----------------------------------+
 column 1: 300 DIP                 column 2: Auto, only while         column 3: * (destinations above
 command bar, search, actions,     Options or Journal is open,        the results; at least 4 rows at
 status                            vertical ScrollViewer              a 163 DIP view height)
```

Height check at 163 DIP (150 %, default 300 px):

- Column 1 needs about 132 DIP (28 + 30 + 28 + 18 plus gaps and margins).
- Column 3 keeps about 111 DIP for the results with one line of chips: 7
  rows at today's 15.5 DIP row height, or 4–5 rows at the proposed 24 DIP
  minimum (C15), against 0–5 DIP today. The creation-time height floor (C16)
  adds about 7 DIP at 150 %.
- Column 2 holds the six Options (about 140 DIP once wrapped) inside a
  `ScrollViewer`.

The header no longer floats 1000 DIP to the right: it sits next to the search
box. If the user declines this change, the minimum fix is to make the existing
bar a `WrapPanel` and move the chips to the right of the results. That keeps
the list visible but stays cramped.

### 5.5 Comparison

| Criterion | Proposal A (Reflow) | Proposal B (Command bar) |
|---|---|---|
| Nothing clipped at 200 DIP, EN/FR/DE | yes | yes (no-clip floor 160 / 190 / 190) |
| Header lines at 200 DIP | 2 (wrapped items) | 1 row (two-line label in FR/DE) |
| Header lines at 229 (150 % default) | 2 | 1 |
| Always-visible commands | 7 | 4 + "⋯" (Settings and Help need 2 clicks) |
| Index status for keyboard and screen readers | needs a focusable wrapper | text notice, live region, menu line |
| Targets | unchanged unless resized | 28 DIP header and Undo, 24 DIP rows |
| Rows saved in vertical mode | 0 | 2 (selection header, Journal toggle row) |
| Change size and risk | small, low | medium: new ToggleButton and ContextMenu styles, re-placement triggers |
| Fixes F1, F4, F5, F6, F7, F10 | partly (F1, F5) | yes |

**Recommendation: Proposal B.** It is the only one of the two that keeps a
single uncluttered command row from the target width up, in all three
languages, and it gives the status a text alternative. It also shortens the
pane in Left/Right docking and reuses every existing view-model action. Its
risk sits in the view layer and the theme dictionary only.

---

## 6. Prioritized change list

Priorities: **P1** = must be in A2; **P2** = should; **P3** = could. Quick
wins (QW) can ship before, and independently of, the A/B decision.

| ID | Change | Tag | Fixes | Prio |
|---|---|---|---|---|
| C1 | Options checkbox labels wrap (`TextBlock` content, `TextWrapping="Wrap"`) | QW | F17 | P1 |
| C2 | Auto-class label: `TextWrapping="WrapWithOverflow"`, no `MaxWidth` | QW | F1, F18 | P1 |
| C3 | Classify/Undo: `Grid` (`*`, `Auto`) or `WrapPanel`; one shared instance for both modes | QW | F19 | P1 |
| C4 | Remove the right-docked header `DockPanel`; children in visual order | QW (A) / part of C9 (B) | F22, F24 | P1 |
| C5 | `AutomationProperties.Name` on icon buttons, search box, result rows and row checkboxes | QW | F8 | P1 |
| C6 | Contrast: dark accent `#115EA3`, update-banner text `#201F1E`, input-border brush, Settings `MutedText` | QW | F27–F31 | P1 |
| C7 | Wrapping labels in the Sent-items triage dialog (plus `SizeToContent="Height"`) and on the pinned-banner Delete button | QW | F32, 2.5 | P1 |
| C8 | Search placeholder and accessible name | QW | F2 | P2 |
| C9 | Command bar with a static "⋯" overflow (Proposal B) | Structural | F6, F22, F1 | P1 if B |
| C10 | Index status inside the search box, plus a notice line (live region) and a menu line | Structural | F7, F9, F30 | P2 |
| C11 | Options/Journal as `ToggleButton`s with chevrons; themed `ToggleButton`, `ContextMenu`, `MenuItem` styles with the focus ring | Structural | F5, F26 | P2 |
| C12 | Segoe UI 12 at the view root; remove the `ElementHost` `Font` mapping | Structural | F15, F11 | P2 (ship with C1–C4, never alone: section 2.7) |
| C13 | Segoe MDL2 Assets icons | Structural | F10 | P2 |
| C14 | Top/Bottom: three-column layout, new switch rule | Structural | 2.8 | P1 |
| C15 | Targets: rows at least 24, row buttons 24×24, chip `✕` 20×20, header icons 28×28 | Structural | F12, F13 | P2 |
| C16 | Creation-time, DPI-aware default size (360 px floor; Top/Bottom height floor) and size logging | QW (code) | 2.2, 2.10, 2.11 | P2 |
| C17 | 4 DIP spacing grid; `d:DesignWidth="200"` | QW | F14, F16 | P3 |
| C18 | Chips: `MaxWidth` plus `TextTrimming` on the name | QW | F21 | P2 |
| C19 | Up/Down in the search box move the highlighted result (focus stays in the box) | Structural (small) | F25 | P3 |
| C20 | `PathTrim` reserve computed from the actual button block | QW | F20 | P3 |

---

## 7. Proposed A2 specification

This is written to replace the A2 text in `ROADMAP.md` once the user has
validated the decisions in section 8. It assumes Proposal B, Segoe UI 12 with
MDL2 icons, the three-column horizontal layout and a 200 DIP target. If the
user chooses differently, items 3, 6–7 and 11 change accordingly.

> **A2.** Pane redesign: command bar with overflow, reflowing rows,
> three-column Top/Bottom layout (requires: A1). Files:
> - `Views/MainPaneView.xaml` and `.xaml.cs`;
> - `Views/Themes/Controls.xaml` and `Brushes.{Light,Dark}.xaml`;
> - `Resources/Strings{,.fr,.de}.resx`;
> - `RblClassTaskPaneHost.cs` (font mapping);
> - `RblClassAddIn.cs` (`ApplyDockSize` only, at creation);
> - `ViewModels/MainPaneViewModel.cs` (view-state properties only);
> - `Views/SentItemTriageWindow.xaml` and `SettingsWindow.xaml` (labels).
>
> No new dependency, no Core change.
>
> 1. **Width contract.** At a view width of **200 DIP and above** (184 DIP
>    column), no control is clipped and no label breaks mid-word, in
>    EN/FR/DE, light and dark. Below 200 DIP, rows keep reflowing: no
>    right-docked `DockPanel` chains anywhere in the pane. Set
>    `d:DesignWidth="200"`.
> 2. **Font and icons.** Set `FontFamily="Segoe UI" FontSize="12"` on the
>    `MainPaneView` root. In `RblClassTaskPaneHost`, call
>    `_elementHost.PropertyMap.Remove("Font")` right after creating the
>    `ElementHost`. Secondary text (feedback, journal, banner subject) uses
>    11 DIP and `MutedText`. Icons: `FontFamily="Segoe MDL2 Assets"`,
>    `FontSize="16"`:
>
>    | Command | Glyph |
>    |---|---|
>    | Save | `&#xE74E;` |
>    | Attach | `&#xE723;` |
>    | More | `&#xE712;` |
>    | Refresh | `&#xE72C;` |
>    | Settings | `&#xE713;` |
>    | Help | `&#xE897;` |
>    | Undo | `&#xE7A7;` |
>    | Open folder (eye, keeps v2.9.0.0 A1) | `&#xE890;` |
>    | Path display | `&#xE740;` |
>    | Remove chip | `&#xE711;` |
>    | Clear all | `&#xE8E6;` |
>    | Chevrons | `&#xE76C;` / `&#xE70D;` |
>
>    The easter-egg pig image is kept.
> 3. **Command bar** (replaces `:88-144`). A `Grid` with columns `*, Auto,
>    Auto, Auto`:
>    - column 0: the Auto-class button, `HorizontalAlignment="Left"`,
>      `Padding="8,4"`. Its label is `TextWrapping="WrapWithOverflow"` with no
>      `MaxWidth`, and `AutomationProperties.Name` is
>      `MainPane_AutoClass_Button`;
>    - columns 1–3: Save as, Manage attachments and More, each 28×28 with
>      `Margin="4,0,0,0"`, keeping their tooltips.
>
>    More opens a `ContextMenu` (placement bottom, right-aligned) with Update
>    folders, Settings…, Online documentation, a separator and a
>    non-interactive line bound to `IndexStatusToolTip`. Each menu item calls
>    the existing view-model method, then `ReturnFocusToSearch()`, exactly
>    like today's click handlers. Declare children in visual order.
> 4. **Search row.**
>    - The existing `TextBox x:Name="QueryBox"` keeps every binding and
>      handler, with `Padding="4,3,20,3"` and
>      `AutomationProperties.Name=MainPane_QueryBox_Name`.
>    - A placeholder `TextBlock` overlays it: `IsHitTestVisible="False"`,
>      `MutedText`, trimmed, visible only while the box is empty.
>    - An 8×8 status ellipse sits inside the box on the right
>      (`Margin="0,0,8,0"`, same colour triggers, faint `Theme.Foreground`
>      outline), with tooltip `IndexStatusToolTip` and name
>      `MainPane_IndexStatus_Name`.
>    - Under the box, a `MutedText` notice line shows `IndexStatusToolTip`
>      only while the index is not ready, with
>      `AutomationProperties.LiveSetting="Polite"`; raise
>      `AutomationEvents.LiveRegionChanged` from its peer when the text
>      changes. Add an `IsIndexReady` view-model property, raised with
>      `IndexStatus`.
> 5. **Results list.** Behaviour unchanged.
>    - The item container style gets `AutomationProperties.Name="{Binding
>      DisplayPath}"`; so does the row `CheckBox`.
>    - Rows get `MinHeight="24"`.
>    - The row buttons become 24×24 MDL2 icons with
>      `MainPane_OpenFolder_Name` and `MainPane_ExpandPath_Name` (tooltips
>      kept).
>    - `v:PathTrim.RightReserve` matches the new button block.
> 6. **Destinations.**
>    - Remove the "Selected: N folders" header row (`:230-236`). Use
>      `SelectionHeader` as the chips container's `AutomationProperties.Name`.
>    - The chips container is a `Grid` with columns `*, Auto`: the chip
>      `WrapPanel` in column 0, and a 24×24 clear-all icon (`&#xE8E6;`) in
>      column 1, top-aligned, with the existing tooltip and name
>      `MainPane_ClearAll_Name`.
>    - Chip names get `MaxWidth="150"` and `TextTrimming="CharacterEllipsis"`
>      (full-path tooltip kept).
>    - The remove buttons are 20×20 with name
>      `MainPane_RemoveDestination_ToolTip`.
> 7. **Toggles row.** A `Grid` with columns `*, Auto`:
>    - Options `ToggleButton` (`IsChecked` two-way to `IsOptionsExpanded`,
>      `x:Name="OptionsToggle"`) on the left;
>    - Journal `ToggleButton` (two-way to `IsStatusLogExpanded`, visible when
>      `HasStatusEntries`, `x:Name="JournalToggle"`) on the right.
>
>    Content is a chevron that turns via a trigger, plus `Pane_Options_Label`
>    / `Pane_Journal_Label`. Add implicit `ToggleButton`, `ContextMenu`,
>    `MenuItem` and `Separator` styles to `Controls.xaml`, using the theme
>    brushes, with the checked state on `Theme.SelectionBackground` and the
>    `Theme.FocusBorder` ring on `IsKeyboardFocused`, in both themes.
> 8. **Options panel.** Each `CheckBox` gets `<TextBlock
>    TextWrapping="Wrap" …/>` content. Keep `x:Name="FirstOptionCheckBox"`
>    and the Tab-from-search behaviour.
> 9. **Action row.** A `Grid` with columns `*, Auto`, a single instance for
>    both modes:
>    - Classify: `AccentButton`, stretched, `Padding="10,5"`, content a
>      centred, wrapping `TextBlock` bound to `ClassifyButtonText`;
>    - Undo: a 28×28 icon, `IsEnabled="{Binding CanUndo}"`, name
>      `MainPane_Undo_Button`, tooltip `MainPane_Undo_ToolTip` with
>      "(Ctrl+Z)" / "(Strg+Z)" appended.
> 10. **Status.** The search feedback gets
>     `AutomationProperties.LiveSetting="Polite"`, with `LiveRegionChanged`
>     raised when `SearchFeedback` changes. The journal is unchanged.
> 11. **Horizontal layout** (replaces v2.9.0.0 A2's bar).
>     - `IsHorizontalLayout = ActualWidth >= 560 && ActualWidth >
>       ActualHeight` (`MainPaneView.xaml.cs:32-35`).
>     - When true, `DataTrigger` setters re-place the same elements into
>       three columns: column 0 (`Width="300"`) holds the command bar, search,
>       action row and status; column 1 (`Auto`, visible only while Options or
>       Journal is open, in a vertical `ScrollViewer`) holds the Options panel
>       and the journal; column 2 (`*`) holds the destinations above the
>       results.
>     - Delete the duplicated horizontal `StackPanel` (`:272-285`) and
>       `CanShowJournalInVerticalMode`.
>     - Check: at a 1200 × 163 DIP view, with one line of chips and Options
>       open, at least 4 result rows (24 DIP each) are visible and nothing is
>       clipped.
> 12. **Contrast.**
>     - Dark `Theme.AccentBackground` `#FF115EA3`, `AccentHover` `#FF0F548C`,
>       `AccentPressed` `#FF0C3B5E`.
>     - Update banner text `#FF201F1E` on `#FF8C00`.
>     - A new `Theme.InputBorder` brush (light `#FF8A8886`, dark `#FFA19F9D`)
>       for the `TextBox` and `ListBox` borders.
>     - `SettingsWindow.xaml:100` and `:112` use `MutedText` instead of
>       `#777`.
> 13. **Banners and dialogs.**
>     - The pinned-banner Delete button content wraps.
>     - The Sent-items triage buttons get wrapping `TextBlock` content and
>       the window gets `SizeToContent="Height"` (width 380, still not
>       resizable).
> 14. **Task-pane size (creation only).** `ApplyDockSize`, still called only
>     from `EnsureTaskPane` before `Visible` and before the
>     `DockPositionStateChange` subscription, sets:
>     - Left/Right: `Width = Math.Max(360, ceil(229 × s) + 16)`;
>     - Top/Bottom: `Height = Math.Max(300, ceil(207 × s))`;
>
>     where `s = GetDpiForWindow(hwnd) / 96.0` and `hwnd` is the Outlook main
>     window (`Process.GetCurrentProcess().MainWindowHandle`, as
>     `DialogPlacement` uses; `GetDpiForWindow` requires Windows 10 1607,
>     below the 1809 minimum). Log the requested size and `s`, and log the
>     host's pixel size on its first `OnResize`. No size is ever set after
>     creation.
> 15. **Invariants (acceptance):**
>     - the `/` jump;
>     - focus return after every pane action, including the More menu items
>       and the toggles;
>     - Ctrl+Z, Enter / Ctrl+Enter, Ctrl tap, Tab to Options;
>     - the focus ring on every focusable element in both themes;
>     - a localized tooltip and `AutomationProperties.Name` on every
>       icon-only button;
>     - the CTP never resized after creation.
> 16. **Strings.** Add these keys to all three resx files, remove
>     `Pane_Options` and `Pane_Journal`, and keep `Status_SelectedFolders_*`
>     (now an accessible name):
>
>     | Key | EN | FR | DE |
>     |---|---|---|---|
>     | `MainPane_More_Name` | More actions | Plus d'actions | Weitere Aktionen |
>     | `MainPane_More_ToolTip` | More: update folders, settings, online documentation | Plus : actualiser les dossiers, paramètres, documentation en ligne | Mehr: Ordner aktualisieren, Einstellungen, Online-Dokumentation |
>     | `MainPane_UpdateFolders_MenuItem` | Update folders | Actualiser les dossiers | Ordner aktualisieren |
>     | `MainPane_Settings_MenuItem` | Settings… | Paramètres… | Einstellungen… |
>     | `MainPane_Help_MenuItem` | Online documentation | Documentation en ligne | Online-Dokumentation |
>     | `MainPane_SaveAs_Name` | Save as .msg | Enregistrer en .msg | Als .msg speichern |
>     | `MainPane_ManageAttachments_Name` | Manage attachments | Gérer les pièces jointes | Anlagen verwalten |
>     | `MainPane_OpenFolder_Name` | Open folder | Ouvrir le dossier | Ordner öffnen |
>     | `MainPane_ExpandPath_Name` | Change path display | Changer l'affichage du chemin | Pfadanzeige ändern |
>     | `MainPane_ClearAll_Name` | Clear all destinations | Retirer toutes les destinations | Alle Ziele entfernen |
>     | `MainPane_QueryBox_Name` | Search folders | Rechercher des dossiers | Ordner suchen |
>     | `MainPane_QueryBox_Placeholder` | Search folders… | Rechercher un dossier… | Ordner suchen… |
>     | `MainPane_IndexStatus_Name` | Folder index status | État de l'index des dossiers | Status des Ordnerindex |
>     | `Pane_Options_Label` | Options | Options | Optionen |
>     | `Pane_Journal_Label` | Journal | Journal | Journal |
>
>     Also reword `IndexStatus_NotFound` to point at "⋯ > Update folders".
>     This also fixes the EN wording "Refresh folders", which does not match
>     the command's name.
> 17. **Docs** (doc-keeper): update the walkthrough's "Pane icon buttons"
>     section, the shortcuts note (Undo icon) and the pane-position note
>     (sizes).
> 18. **Verification.** A live check with the user at a 200 DIP view (about
>     316 px of frame at 150 %) and at the default width, in EN/FR/DE, light
>     and dark, docked Left/Right and Top/Bottom. Optionally run an offline
>     layout check with the method in Appendix A of `docs/ui-review.md`.

---

## 8. Decisions for the user

1. **Target minimum width.** Recommendation: **200 DIP of view width**
   (about 316 px of frame at 150 %), instead of the 260 DIP starting
   proposal. 260 DIP is wider than today's default at 150 % and would bring
   back the compact message list. 180 DIP gives more headroom but leaves
   about 11 characters of path and two-line buttons in FR/DE.
2. **Layout.** Recommendation: **Proposal B**, the command bar with a "⋯"
   overflow and merged rows. Proposal A keeps all seven commands visible but
   needs a two-line header at the default and target widths.
3. **What goes behind "⋯".** Recommendation: **Update folders, Settings and
   Online documentation**; Auto-class, Save as and Manage attachments stay
   visible. This partly revisits v2.10.0.0 B1 / v2.11.0.0 B2, which put
   Settings and Help in the pane header: they stay in the pane, one click
   deeper. Alternative: also move Save as and Manage attachments, so the
   Auto-class label stays on one line even at 200 DIP.
4. **Top/Bottom docking.** Recommendation: **replace** the one-line action
   bar with the three-column layout, keeping the automatic switch plus a
   560 DIP minimum. At 150 % today the results list disappears as soon as a
   destination is checked. Alternatives: keep the bar (minimal fix: chips
   beside the results), or drop the horizontal mode entirely.
5. **Look: font, icons, scaling.** Recommendation: **Segoe UI 12 plus Segoe
   MDL2 Assets icons, no proportional font scaling.** This matches Outlook,
   gives consistent icons, and keeps text at a readable fixed size.
   Alternative: keep today's (accidental) Microsoft Sans Serif 11 and emoji.
   Narrower, but foreign-looking and inconsistent.

---

## Appendix A: measurement method

- **Live pane (read-only).** From a separate PowerShell process running
  per-monitor DPI aware:
  - `EnumChildWindows` and `GetWindowRect` on the Outlook frame found the
    `MsoCommandBar "RBLclass"` frame (360 px), the `MsoWorkPane` (344 px) and
    the WPF `HwndWrapper` host (344 × 1183 px);
  - `AutomationElement.FromHandle` on the host listed every control's name,
    help text and bounding rectangle.

  No input, focus, resize, redock, close or reload was performed. Outlook
  was minimized, so `PrintWindow` returned no image and the window was left
  as it was.
- **Screenshots.** Luminance-edge scans along rows through the search box
  and the header of the two #67 images (1912 × 265 and 1909 × 298 px at
  144 dpi).
- **Offline harness.** A PowerShell STA process (never Outlook):
  1. load the Release `RBLclass.AddIn.dll` and its `Serilog` and
     `RBLclass.Core` dependencies with `Assembly.LoadFrom`;
  2. set `TaskPaneServices.Localization` to `new LocalizationService("en" |
     "fr" | "de")`;
  3. create `MainPaneView`;
  4. merge `Controls.xaml` and `Brushes.*.xaml`. In PowerShell, use
     `ResourceDictionary.set_Source(...)`: `$d.Source = …` adds a dictionary
     *key* instead;
  5. set a stub DataContext exposing the bound property names;
  6. reproduce the live font with `FontFamily = "Microsoft Sans Serif"`,
     `FontSize = 11` on the view;
  7. `Measure` / `Arrange` at width *W*;
  8. flag a control when its layout slot is narrower than its natural width
     (measured at *W* = 4000), or when its bounds leave the column;
  9. binary-search the smallest *W* per language and state.

  Results are calibrated against the live pane in section 2.3. Proposal B's
  prototype uses the same pipeline on loose XAML, with a hyphen-aware "no
  mid-word break" check and a downward scan. The scripts lived in the session
  scratchpad and are not committed. Turning them into a `tools/` script is
  optional A2 support.
- **Contrast.** WCAG relative-luminance formula on the brush values in
  `Brushes.Light.xaml:5-26` and `Brushes.Dark.xaml:5-26` and the colours
  hard-coded in `MainPaneView.xaml` and `SettingsWindow.xaml`.

## Appendix B: natural widths today

All values are in DIP, measured with the live font. Each value includes the
control's own margin.

| Element | EN | FR | DE |
|---|---|---|---|
| Classify, nothing checked ("Classify" / "Classer" / "Klassieren") | 64.6 | 63.3 | 77.5 |
| Classify, 2 destinations (SemiBold) | 125.7 | 147.5 | 138.1 |
| Classify, 12 destinations (SemiBold) | 132.1 | 153.8 | 144.5 |
| Undo (incl. 6 DIP left margin) | 58.3 | 70.0 | 90.2 |
| Options toggle | 60.4 | 60.4 | 67.1 |
| Journal toggle | 58.5 | 58.5 | 58.5 |
| Selection header, 12 destinations | 101.0 | 133.3 | 116.5 |
| Clear button | 37.9 | 47.6 | 44.6 |
| Option: list every folder | 143.1 | 208.4 | 178.0 |
| Option: keep a copy | 189.7 | 279.0 | 213.0 |
| Option: remove attachments | 214.3 | 251.7 | 239.5 |
| Option: whole conversation | 168.8 | 155.4 | 191.0 |
| Option: strip banner | 175.5 | 228.3 | 208.2 |
| Option: mark as read | 81.9 | 109.5 | 126.7 |
| Pinned "Delete from Inbox" | 100.6 | 181.7 | 138.0 |
| Update banner "Dismiss" | 46.6 | 42.9 | 58.1 |

For comparison, the same Segoe UI 12 strings that Proposal B relies on:

| String | EN | FR | DE |
|---|---|---|---|
| Auto-class label | 55.2 | 87.7 | 90.6 |
| Longest unbreakable part of the label | 30.4 | 60.4 | 60.2 |
| "Classify to 12 folders" (SemiBold) | 111.3 | 129.0 | 123.6 |
| Search placeholder | 83.6 | 125.3 | 86.9 |

## Sources

- GitHub issue #67: <https://github.com/reboulip/rblclass/issues/67>
- Microsoft Q&A, compact-layout setting:
  [How to restore Default Outlook view](https://learn.microsoft.com/en-us/answers/questions/4663823/how-to-restore-default-outlook-view)
- ExtendOffice, Outlook view "Other Settings":
  [Other settings](https://www.extendoffice.com/outlook-receive-read-emails/outlook-inbox/change-inbox-view/other-settings.html)
- WCAG 2.2, success criteria 1.4.1, 1.4.3, 1.4.4, 1.4.11, 2.4.3, 2.4.7,
  2.5.8 and 4.1.2: <https://www.w3.org/TR/WCAG22/>
