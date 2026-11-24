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
    Outlook assigns it for that session. The standard size is restored the
    next time Outlook starts: at least 360 px wide for Left/Right docking and
    at least 300 px tall for Top/Bottom docking, enlarged on high-DPI displays
    so the content always fits. It is applied only when the pane is created, so
    you can resize the pane freely with the splitter at any time.

!!! note "Three-column layout"
    When the pane is docked Top or Bottom and is at least 560 DIP wide and
    wider than it is tall, it switches to three columns: the command bar, search
    box and action buttons on the left, the **Options** and **Journal** panel in
    the middle, and the destination chips above the results on the right.

## The ribbon tab

The **RBLclass** ribbon tab has two buttons:

- **RBLclass pane** — show or hide the pane.
- **About** — version, author and technical information, plus an **Online
  documentation** link that opens this site in your default browser.

**Settings** and **Update folders** have moved into the pane itself (see
[the command bar](#pane-icon-buttons) below).

!!! tip "Keyboard access to the pane"
    Press ++alt++, then ++r++ (the RBLclass tab), then ++r++ again (the pane
    button).

## Pane icon buttons

The top row of the task pane is a command bar:

- **Auto-class**, **Save as file**, and **Manage attachments** buttons (described
  under [Classify](#classify), [Save as file](#save-as-file) and
  [Attachment management](#attachment-management)).
- **⋯** (More) — a menu holding:
    - **Update folders** — re-scans your mailboxes so folders you created or
      renamed directly in Outlook appear in search.
    - **Settings…** — opens the Settings dialog (see the
      [Settings reference](#settings-reference)).
    - **Online documentation** — opens this documentation site in your default
      browser.
    - The index status line.

The search box shows a placeholder, with the index-status dot inside it: yellow
while a scan runs, green when it finishes. If the index has not been built yet,
the pane points you to **⋯ → Update folders**.

## Features

### Folder search

Type keywords in the search box at the top of the pane. Results appear after a
short pause. For each result row:

- **Single-click** the row to check it as a filing destination (check several to
  file into all of them at once).
- ++enter++ files the selected mail into the highlighted folder.
- ++ctrl+enter++ *opens* the highlighted folder in Outlook instead of filing.
- **Double-click** a row to file into that folder immediately.
- The **eye** icon opens that folder in Outlook.
- Right-click → **Add subfolder…** creates a sub-folder under it (with a targeted
  re-index so it's searchable straight away).
- Right-click → **Rename folder…** opens a small dialog pre-filled with the
  folder's current name; on confirm, the folder is renamed in the live Outlook
  store and the index is refreshed. This item is grayed out on a store's root
  folder (the top-level PST entry) — store roots cannot be renamed. Any other
  folder or sub-folder can be renamed. If Outlook refuses the name (for example,
  a duplicate or a reserved name), a status message is shown instead of a crash.
- Right-click → **Delete folder** removes the folder from the Outlook store. A
  confirmation dialog first tells you how many mail items and direct sub-folders
  the folder contains; clicking **Yes** proceeds even when the folder is
  non-empty. Outlook's soft-delete semantics apply: the folder is moved to
  *Deleted Items* unless it is already inside *Deleted Items*, in which case it
  is permanently deleted. This item is grayed out on a store's root folder, the
  same restriction as **Rename folder**. On failure, a status message is shown in
  the pane instead of a crash.
- The **path-display** icon cycles how that row shows its path (see *Search
  result path display* in Settings).

Checked destinations appear as chips above the results, with a clear-all icon to
uncheck them at once.

Tapping ++ctrl++ on its own (press and release, without another key) toggles
*"List every matching folder"*.

### Classify

Select one or more mails in Outlook, check the destination folder(s) in the
pane, then click **Classify** (or press ++enter++ / double-click for a single
folder).

**Options** and **Journal** are collapsible panels toggled with a chevron. The
**Options** zone offers five per-action overrides, seeded from your settings
and **reset after each classify**:

- Keep a copy
- Remove attachments from the filed copy
- Include the whole conversation
- Strip the external-sender banner
- Mark as read

**Auto-class** files the selected mail(s) to wherever their conversation was last
classified, without picking a folder by hand.

### Undo

The **Undo** icon next to the Classify button — or ++ctrl+z++ anywhere in the pane — reverses the last
classify: moved mails go back to their source folder (and are re-selected there
so you can re-file them elsewhere) and filed copies are removed.

!!! warning "Undo has limits"
    Only the **last** classify can be undone, and **removed attachments cannot be
    restored**.

### Save as file

The **Save as file** button writes the selected mail(s) to a folder you pick, as
Outlook messages (**.msg**) or as **PDF**, depending on the *Save as format*
setting (see the [Settings reference](#settings-reference)). Only mails are saved.

PDF export converts each mail with **Microsoft Word**, which must be installed;
Word runs in the background and is not shown. The busy overlay reads *"Saving as
PDF… n of N"*. Each PDF is named after the mail's subject (made file-name safe),
with " (1)", " (2)"… added when the name already exists in the folder. If Word is
not available, nothing is written and the status line reads *"PDF export needs
Microsoft Word"*.

!!! warning "Encrypted mail in a PDF"
    If any selected mail is encrypted, one confirmation covers the whole batch: the
    PDF would contain those mails **unencrypted**. Choose **Continue** to include
    them, or **Skip encrypted mails** (the default, also ++esc++) to save only the
    others. Skipped mails are listed in the pane journal, separately from failures.

### Attachment management

When *"Remove attachments"* is on, the filed copy loses its attachments; the
original (or the kept copy) is never touched. In the default *"Ask what to do
with each attachment"* mode (and from the pane's **Manage attachments** button),
a dialog lets you **Delete**, **Save to…** or **Keep** attachments. Radio buttons
at the top choose one of three modes:

1. **Same action and folder for all attachments** (default) — one choice and one
   folder applied to every attachment; the list shows each attachment with its
   size.
2. **An action per attachment, one folder for the saved ones** — each attachment
   has its own Delete / Save to / Keep, and one folder at the top is used by
   every attachment set to Save to.
3. **An action and a folder per attachment** — each attachment has its own
   action and its own folder.

Switching back to the first mode never overrides an attachment you explicitly set
to Delete or Keep in the other two; it is shown with its own action in the list.
If you set a *Default folder for saved attachments* in Settings, it pre-fills the
save folder.

If *Offer the images of the message body for removal* is on in Settings (off by default), the window
also lists images shown inside the message body (`cid:` inline images). They carry
a **body image** badge, default to **Keep**, and a notice explains that they are
part of the message. Embedded items, OLE objects and remote `http` images are
never listed. These images are offered even when *Remove attachments* is off; in
that case only the images are listed and regular attachments stay untouched, and
the **Manage attachments** button also works on a mail whose only attachments are
body images. Removing one rewrites the message's HTML body, replacing the image
with a text placeholder ("[Image removed]"), and then deletes the attachment; it
is listed in the disposition label like any other removed attachment. **Save
to…** works for body images too. Body images are never removed silently: without
the window, the "remove attachments" strip still skips inline images, so
signature logos survive.

!!! danger "Encrypted mail is never stripped"
    Encrypted (S/MIME) mails are never stripped — RBLclass skips attachment
    removal for them and reports it, so a filed copy never silently loses
    protected content. This holds even when a mail-encryption product decrypts
    the open mail in memory: RBLclass also checks the class stored in the folder,
    the S/MIME payload and whether the mail was already modified when opened.
    The filed copy is then left untouched. The same applies to the other changes
    RBLclass would make to a mail: encrypted or signed mails are never rewritten,
    never marked as read, and a flagged one stays flagged and is simply not
    offered in the "mark flagged tasks complete" prompt. The experimental
    re-encrypted filing below is the only exception.

    Filing, copying, Undo and deleting an encrypted mail work as usual, and the
    mail **stays encrypted** in its new folder: RBLclass moves it as-is, without
    opening it, so a mail-encryption product that has decrypted it in memory
    (smart card inserted) cannot leak that decrypted version into the filed
    mail. If the filed mail cannot be moved, it stays where it was, still
    encrypted, and counts as an error. In the rare case where its new position
    cannot be identified, the mail is filed but cannot be undone.

### Re-encrypted filing (experimental)

!!! warning "Experimental and not undoable"
    This feature is for pilot tests and is **off by default** (Settings →
    **Experimental**). A re-encrypted filing **cannot be undone**: there is no
    Undo entry, and Undo does not roll back its Auto-class history.

Without this option, an encrypted mail is always filed as it is and stays
encrypted. When **Re-encrypt encrypted mail when filing** is on, it only matters
if the filing would change the mail: you remove its attachments, strip the
external-sender banner, or delete or save out an attachment or body image in the
attachment window. A plain classify still files the mail as it is. In that case,
Classify and Auto-class handle an **encrypted** mail that your mail-encryption
product has decrypted in memory (smart card inserted) differently:

1. The mail is rebuilt from its decrypted content (body, body images,
   attachments).
2. Your classify options are applied to that content: remove attachments,
   per-attachment choices (including **Save to…**), body images, the "former
   attachments" label and the external-sender banner strip.
3. The result is encrypted again (S/MIME) to your certificate and stored as an
   encrypted mail in every destination, with the original subject, sender,
   recipients, dates, categories, importance and flag. It keeps the original's
   read state unless **Mark as read** is on.
4. Unless **Keep a copy** is on, the original goes to its store's Deleted Items.
   If Outlook wrote that original decrypted, it is deleted permanently instead.
   Mail already in Deleted Items is not moved.

Every filed copy is checked as stored encrypted before the original is touched.
In the attachment window, an eligible encrypted mail is offered like a plain
mail; **Manage attachments** (in place) still never offers protected mail.

The status line reads *"N encrypted mail(s) re-encrypted (cannot be undone)."*

**Not applicable** (standard behaviour, silently): signed-only mail, mail still
showing its encrypted part (smart card absent, so not decrypted), and mail
protected only by "unsaved changes".

**Fallbacks.** If a mail cannot be re-encrypted, it is filed the standard way and
the status line adds *"N encrypted mail(s) filed without re-encryption — see the
log"*. Causes are: no usable certificate (a one-time message, *"No usable
encryption certificate: re-encryption skipped."*, is also shown per Outlook
session); content that cannot be rebuilt (embedded item or `.msg`, OLE or
by-reference attachment, attachments over 25 MB in total, a read failure); an
encryption failure; or a copy that cannot be created or verified (any copy
already made for that mail is removed first).

**Limits.** No undo; the sender's signature is not kept; the smart card must be
inserted so the mail is readable; the original lands in Deleted Items (or is
permanently deleted when written decrypted); 25 MB attachment cap.

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
and forwards, and from filed copies. A received mail is never rewritten: the
strip only touches the reply or forward you are composing.

### Phishing warnings

Once a banner is learned, RBLclass also uses it to recognise external mail and
asks for confirmation before something in it is opened:

- **Link** — clicking a link, pressing ++enter++ on a selected link, or choosing
  **Open hyperlink**. In a reply, ++ctrl+click++ is covered; a plain click does
  not follow links, so it does not warn. `mailto:` links and `#` anchors never
  warn.
- **Attachment** — opening or previewing one. After **Open anyway**, the same
  mail does not ask again for 4 seconds.

A **Possible phishing** dialog shows the sender's display name and real SMTP
address, plus the link destination (and, when the mail displays something
different, a *"Shown in the email as: …"* line) or the attachment's file name.
**Open anyway** continues (Outlook's own security prompts still apply);
**Cancel**, the default and ++esc++, opens nothing. Without a learned banner it
never warns, and if the check itself fails the item opens.

!!! note "After Cancel in the reading pane"
    Outlook leaves the reading pane on *"Loading…"* with its own *"Back to
    message"* link; click it to return to the mail.

!!! warning "Touch taps are not covered"
    The link warning cannot block a link opened by a touch tap. This is a known,
    accepted gap; mouse and keyboard are covered.

## Settings reference

Open **Settings** from the **⋯** menu in the command bar of the RBLclass pane
(**Settings…**). Every change applies and persists immediately — there is no
"Apply" button, and closing is always safe.

### Folder search

| Setting | Default | What it does |
| --- | --- | --- |
| List every matching folder | Off | When off, duplicate branches of the tree are collapsed; when on, every matching folder is listed separately. |
| Match only the start of words | Off | Stricter matching: "proj" matches "Projects" but "security" does *not* match "Cybersecurity". The default (off) is the broader "contains" search. |
| Open folders in a new Outlook window | Off | When on, "open folder" opens a new Outlook window instead of navigating the current one. |
| Search result path display | Single line | How each result row shows its folder path: a single left-trimmed line, the full path wrapped to the pane width, or the full path with one sub-folder per line. The per-row path-display icon cycles all three for that row. |
| Maximum results to display | 50 | Cap on the number of result rows. Applied after the depth filter below. |
| Maximum search depth (0 = unlimited) | 0 | Limits how deep in the folder tree a folder can match. The store root is level 0 and its direct sub-folders are level 1. Folders deeper than N are not match candidates; they only appear inside a matching ancestor's collapsed `\| [...]` entry. With *List every matching folder* on, folders deeper than N are not listed, and a level-N folder that has sub-folders gets a ` \| [...]` suffix. 0 means unlimited; negative values become 0; the maximum is 100. Favourite filesystem folders are not affected. |
| Search after at least N characters | 2 | Avoids searching on a single letter. Range 1–10. |
| Pause before searching (ms) | 200 | Search fires this many milliseconds after your last keystroke. Range 0–2000. |

### Classify

| Setting | Default | What it does |
| --- | --- | --- |
| Keep the original after classifying | Off | Copies the mail to the destination and leaves the original in place ("Keep a copy"). |
| Remove attachments from the filed copy | Off | Strips attachments from the copy that lands in the destination; the original is untouched. Sub-options: ask per attachment (default) or delete all silently; and whether to record a "former attachments" note in the body. |
| Default folder for saved attachments (Settings → Attachments) | *Not configured* | **Pick folder…** / **Clear**. Pre-fills the save folder in the attachment window. If the folder cannot be reached (for example a disconnected network share), the field is simply left empty; the window still opens immediately. |
| Offer the images of the message body for removal (Settings → Manage attachments) | Off | Also lists body (inline `cid:`) images in the attachment window so you can remove or save them; they default to Keep. Removing one rewrites the HTML body (placeholder text) and cannot be undone. Never applied to encrypted mail. |
| Save as format (Settings → Attachments → Manage attachments) | Outlook message (.msg) | What the pane's **Save as file** button writes: **Outlook message (.msg)**, **PDF (.pdf)** or **Ask each time**. With *Ask each time*, the folder picker shows a format choice with .msg preselected; that choice applies to that action only and is not saved. PDF export needs Microsoft Word installed — see [Save as file](#save-as-file). |
| Keep a safety copy in Deleted Items | Off | When the original is not kept, also places a copy in Deleted Items as a safety net. |
| Include the whole conversation by default | Off | Pre-checks the pane's "Include the whole conversation" option. |
| Classify meeting requests and responses | Off | Allows meeting items selected in Outlook to be classified alongside regular mails. |
| Mark classified mail as read | Off | Marks each filed mail as read as part of classify (including Auto-class). |
| Auto-class history window (days) | 90 | How far back Auto-class looks when deciding where a conversation was last filed. Range 1–3650. |

### External-sender banner

*Learn from selected mail* captures the banner from the currently selected mail;
*Forget* clears it; *Diagnose* reports whether the current mail matches;
*Save diagnostic file* writes a self-contained HTML file to
`%LocalAppData%\RBLclass\diagnostics\` — the mail body with the matched banner
highlighted — useful for sending to your support contact when diagnosing why a
banner is not stripped correctly. Then five toggles, only useful once a banner is
learned:

| Setting | Default | What it does |
| --- | --- | --- |
| Strip on reply/forward, on classify, during Auto-class | Off | Remove the banner in each case (see [External-sender banner strip](#external-sender-banner-strip)). |
| Warn before opening a link from an external email | On | Confirmation before a link opens (see [Phishing warnings](#phishing-warnings)). |
| Warn before opening an attachment from an external email | On | Confirmation before an attachment opens or previews. |

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

### Experimental

Features for pilot tests; a warning at the top of the tab says so.

| Setting | Default | What it does |
| --- | --- | --- |
| Re-encrypt encrypted mail when filing | Off | Encrypted mail is always filed as it is and stays encrypted. When you also remove attachments or the external banner (or delete or save out an attachment or body image) from a readable encrypted mail, Classify and Auto-class rebuild it without them, encrypt it again and file it as an encrypted copy. Not undoable — see [Re-encrypted filing](#re-encrypted-filing-experimental). |
| Re-encryption certificate | *Automatic* | The certificate used to encrypt the filed copy. The list shows **Automatic** plus each candidate as "Subject (expires yyyy-MM-dd)", flagged *private key not visible* when Windows cannot see its private key, and an *Unavailable certificate* entry if the stored one disappeared. A line under the list says which certificate Automatic resolves to, or that none is usable. |

Candidates come from your personal certificate store (Current User → Personal):
valid now (dates only, no revocation check, no network access), RSA key, allowed
for key encipherment and for secure e-mail. **Automatic** also requires that the
certificate's e-mail address equals your default account's address and that
Windows sees its private key; the latest expiry wins. An explicit choice may be any
candidate; if it stops being valid it is not silently replaced and re-encryption
is skipped.

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
| ++slash++ | Move focus to the search box from anywhere else in the pane (the `/` is not typed). Inside the search box it types `/` normally. Pane only — it does not work from Outlook's message list. |

After a pane action or a dialog closed from the pane (classify, Auto-class,
Undo, Save as file, Manage attachments, Settings, Update folders and so on), keyboard
focus returns to the search box with its text selected, so you can type
straight away. It does not take focus while another Outlook window, such as a
compose window, is in front. **Open folder** (the row button and
++ctrl+enter++) is the exception: it gives focus to Outlook's message list.

The control that currently has keyboard focus is shown with a thin blue outline
(lighter blue in the dark theme; white on the blue Classify button), in the pane
and in every RBLclass window. The outline also stays after a mouse click until
focus moves elsewhere.

!!! note "Known issue: Delete key after using the pane"
    After the pane has had keyboard focus, Outlook's ++delete++ key does nothing
    in the message list until you click in the message list once (tracked in
    issue #49). Workaround: click the message list, then press ++delete++.

## Refresh the folder index

The folder index is built automatically in the background when Outlook starts. If
you create or rename a folder directly in Outlook and it doesn't appear in
search, open the **⋯** menu in the command bar of the RBLclass pane and choose
**Update folders**.

## Suggesting improvements / reporting problems

RBLclass is maintained on
[GitHub Issues](https://github.com/reboulip/rblclass/issues). Open an issue there
to report a bug or suggest a feature. If you cannot reach GitHub from your
workstation, use the contact channel from the installation guide instead. To help
diagnose a problem, attach the latest log file from
`%LocalAppData%\RBLclass\logs\`.
