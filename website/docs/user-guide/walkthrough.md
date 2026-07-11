# Complete walkthrough

This page covers every feature and every setting. For a two-minute intro, see the
[Quick start](quick-start.md).

RBLclass adds a **task pane** on the right of Outlook and an **RBLclass** tab on
the ribbon.

## Task pane position

The pane opens docked to the right of the Outlook window by default. You can
drag it to any edge (Left, Right, Top, or Bottom) using Outlook's standard
splitter; the position is remembered across restarts.

!!! note "Pane size after live-redocking"
    If you drag the pane to a new edge during a session, it keeps whatever size
    Outlook assigns it for that session. The standard size (360 px wide for
    Left/Right docking, 300 px tall for Top/Bottom docking) is restored the
    next time Outlook starts. You can resize the pane freely with the splitter
    at any time.

## The ribbon tab

If you close the pane, the **RBLclass** ribbon tab has four buttons:

- **RBLclass pane** — show or hide the pane.
- **Refresh folders** — re-scan your mailboxes so folders you created or renamed
  directly in Outlook show up in search.
- **Settings** — all options (see the [Settings reference](#settings-reference)).
- **About** — version, author and technical information.

!!! tip "Keyboard access to the pane"
    Press ++alt++, then ++r++ (the RBLclass tab), then ++r++ again (the pane
    button).

## Features

### Folder search

Type keywords in the search box at the top of the pane. Results appear after a
short pause. For each result row:

- **Single-click** the row to check it as a filing destination (check several to
  file into all of them at once).
- ++enter++ files the selected mail into the highlighted folder.
- ++ctrl+enter++ *opens* the highlighted folder in Outlook instead of filing.
- **Double-click** a row to file into that folder immediately.
- The **👁** button opens that folder in Outlook.
- Right-click → **Add subfolder…** creates a sub-folder under it (with a targeted
  re-index so it's searchable straight away).
- The **⤢** button cycles how that row shows its path (see *Search result path
  display* in Settings).

Tapping ++ctrl++ on its own (press and release, without another key) toggles
*"List every matching folder"*.

### Classify

Select one or more mails in Outlook, check the destination folder(s) in the
pane, then click **Classify** (or press ++enter++ / double-click for a single
folder).

The **Options** zone offers five per-action overrides, seeded from your settings
and **reset after each classify**:

- Keep a copy
- Remove attachments from the filed copy
- Include the whole conversation
- Strip the external-sender banner
- Mark as read

**Auto-class** files the selected mail(s) to wherever their conversation was last
classified, without picking a folder by hand.

### Undo

The **Undo** button — or ++ctrl+z++ anywhere in the pane — reverses the last
classify: moved mails go back to their source folder (and are re-selected there
so you can re-file them elsewhere) and filed copies are removed.

!!! warning "Undo has limits"
    Only the **last** classify can be undone, and **removed attachments cannot be
    restored**.

### Attachment management

When *"Remove attachments"* is on, the filed copy loses its attachments; the
original (or the kept copy) is never touched. In the default *"Ask what to do
with each attachment"* mode, a dialog lets you **Delete**, **Save to…** or
**Keep** each attachment; a *"use the same save directory"* checkbox applies your
choice to all of them.

!!! danger "Encrypted mail is never stripped"
    Encrypted (S/MIME) mails are never stripped — RBLclass skips attachment
    removal for them and reports it, so a filed copy never silently loses
    protected content.

### Sent-item triage

Each time a mail lands in **Sent Items**, RBLclass can act automatically or ask
you. The choices are:

- **Move to Inbox**
- **Move to archive folder** (a folder you configure)
- **Delete from Sent Items**
- **Leave in Sent Items**

When a sent mail is moved to the Inbox, the pane can auto-open with that mail
pinned as the next thing to file.

### Send-time guards

- **Forgotten attachment** — if your message body mentions an attachment (using
  the configured keywords) but nothing is attached, a confirmation appears
  before the mail goes out.
- **External recipient** — before sending, a confirmation lists any recipient
  whose domain is not on your internal-domains list.

### External-sender banner strip

Many organisations add a *"this email came from outside"* banner to incoming
mail. Teach RBLclass that banner once (**Settings → External-sender banner →
Learn from selected mail**), and it can strip it automatically from your replies
and forwards, and from filed copies.

## Settings reference

Open **Settings** from the RBLclass ribbon group. Every change applies and
persists immediately — there is no "Apply" button, and closing is always safe.

### Folder search

| Setting | Default | What it does |
| --- | --- | --- |
| List every matching folder | Off | When off, duplicate branches of the tree are collapsed; when on, every matching folder is listed separately. |
| Match only the start of words | Off | Stricter matching: "proj" matches "Projects" but "security" does *not* match "Cybersecurity". The default (off) is the broader "contains" search. |
| Open folders in a new Outlook window | Off | When on, "open folder" opens a new Outlook window instead of navigating the current one. |
| Search result path display | Single line | How each result row shows its folder path: a single left-trimmed line, the full path wrapped to the pane width, or the full path with one sub-folder per line. The per-row ⤢ button cycles all three for that row. |
| Maximum results to display | 50 | Cap on the number of result rows. |
| Search after at least N characters | 2 | Avoids searching on a single letter. Range 1–10. |
| Pause before searching (ms) | 200 | Search fires this many milliseconds after your last keystroke. Range 0–2000. |

### Classify

| Setting | Default | What it does |
| --- | --- | --- |
| Keep the original after classifying | Off | Copies the mail to the destination and leaves the original in place ("Keep a copy"). |
| Remove attachments from the filed copy | Off | Strips attachments from the copy that lands in the destination; the original is untouched. Sub-options: ask per attachment (default) or delete all silently; and whether to record a "former attachments" note in the body. |
| Keep a safety copy in Deleted Items | Off | When the original is not kept, also places a copy in Deleted Items as a safety net. |
| Include the whole conversation by default | Off | Pre-checks the pane's "Include the whole conversation" option. |
| Classify meeting requests and responses | Off | Allows meeting items selected in Outlook to be classified alongside regular mails. |
| Mark classified mail as read | Off | Marks each filed mail as read as part of classify (including Auto-class). |
| Auto-class history window (days) | 90 | How far back Auto-class looks when deciding where a conversation was last filed. Range 1–3650. |

### External-sender banner

*Learn from selected mail* captures the banner from the currently selected mail;
*Forget* clears it; *Diagnose* reports whether the current mail matches. Then
three toggles (all off by default, and only useful once a banner is learned):
strip on reply/forward, strip on classify, and strip during Auto-class.

### Sending guards

| Setting | Default | What it does |
| --- | --- | --- |
| Warn before sending outside the organisation | On | Confirmation dialog if any recipient's domain is not on the internal list. |
| Domains treated as internal | *(empty)* | Semicolon-separated domains to exempt from the external-recipient warning, e.g. `contoso.com;subsidiary.fr`. |
| Forgotten-attachment keywords | `attach;enclos;joint;PJ` | If one of these words appears in the body but no attachment is attached, a warning appears before sending. Clear the field to disable the guard. |

### Sent items

| Setting | Default | What it does |
| --- | --- | --- |
| When a mail lands in Sent Items | Ask me each time | Ask each time, or automatically Move to Inbox, Move to archive folder, Delete, or Leave. |
| Archive folder | *(not configured)* | The Outlook folder sent mail is moved to when "Move to archive folder" is chosen. Use *Pick folder…* / *Clear*. |
| Open the classify pane after moving to Inbox | On | Reveals the pane with the sent mail pinned as the next classify target. |

### Updates

| Setting | Default | What it does |
| --- | --- | --- |
| Check for a newer version at startup | Off | When on *and* a URL is set, RBLclass makes one request to that URL at startup to read the latest version number, and shows an orange banner if a newer version is available. No other data is sent, and nothing happens if the URL is blank. |
| Version-check URL | *(empty)* | The address of a plain-text version marker on your internal HTTPS share. Ask your deployment team for this value. |

### Favourite attachment folders

Pre-index Windows directories so the "Save to…" picker (in the attachment
dialog) finds a save location as fast as an Outlook folder. Sub-folders are
indexed to a depth of 4.

### Language

RBLclass follows Outlook's language by default (falling back to English). You can
override it to English, Français or Deutsch.

!!! note
    Changing the language takes effect **after you restart Outlook**.

## Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| ++ctrl+z++ | Undo the last classify (anywhere in the RBLclass pane). |
| ++enter++ | File the selected mail into the highlighted folder. |
| ++ctrl+enter++ | Open the highlighted folder in Outlook. |
| ++ctrl++ (tap) | Toggle "List every matching folder". |
| ++alt++, ++r++, ++r++ | Open / toggle the RBLclass pane from the ribbon. |
| ++tab++ (from the search box) | Open the Options panel and move focus to its first checkbox. |

## Refresh the folder index

The folder index is built automatically in the background when Outlook starts. If
you create or rename a folder directly in Outlook and it doesn't appear in
search, click **Refresh folders** on the ribbon.

## Suggesting improvements / reporting problems

RBLclass is maintained on
[GitHub Issues](https://github.com/reboulip/rblclass/issues). Open an issue there
to report a bug or suggest a feature. If you cannot reach GitHub from your
workstation, use the contact channel from the installation guide instead. To help
diagnose a problem, attach the latest log file from
`%LocalAppData%\RBLclass\logs\`.
