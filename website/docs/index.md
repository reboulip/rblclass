# RBLclass

**RBLclass** is an Outlook add-in that helps you file (classify) emails into
folders across your mailboxes and `.pst` archives faster: type a few keywords to
jump to any destination folder, file the selected mail(s) into one or several
folders with one click, clean up attachments on the way, and get a few
send-time reminders (forgotten attachments, external recipients).

It runs **entirely on your machine** against your local PST files — no server, no
Exchange Online dependency — and installs **per-user with no administrator
rights**. It is a from-scratch C#/.NET rewrite of a legacy Outlook VBA macro of
the same name.

<div class="grid cards" markdown>

- :material-account: **Using RBLclass**

    Install it, file your first email in two minutes, and learn every feature
    and setting.

    [:octicons-arrow-right-24: User Guide](user-guide/installation.md)

- :material-code-braces: **Working on RBLclass**

    The stack, the strictly-layered architecture, how the code is tested, and
    how a build reaches a workstation.

    [:octicons-arrow-right-24: Developer Guide](developer-guide/tech-stack.md)

- :material-shield-check: **Security model**

    Trust boundaries, what data is handled, network posture, and what is
    accepted by design.

    [:octicons-arrow-right-24: Security](security.md)

</div>

## What it does

- **Folder search** — type a few keywords to jump straight to a folder anywhere
  in your PST tree, with collapse/expand and a configurable match mode.
- **One-click classify** — file the selected mail(s) into one or several
  destination folders at once, with per-destination copies and attachment
  handling on copy.
- **Attachment management** — strip attachments from a filed copy without
  touching the original (or a kept copy). Encrypted (S/MIME) mail is never
  stripped.
- **Inline sub-folder creation** — create a destination folder on the fly while
  classifying, with a targeted re-index.
- **Conversation widening** — pull the rest of a conversation into view, with a
  task-completion guard.
- **Send-time guards** — configurable reminders for forgotten attachments
  (keyword list) and external recipients (internal-domain allowlist), wired to
  Outlook's `ItemSend`.
- **Sent-item triage** — a prompt to Class / Delete / Move-to-Inbox / Leave a
  sent item or its whole conversation.
- **Settings** — every option persists across restarts, backed by a local
  SQLite database under `%LocalAppData%\RBLclass\`.

Once installed, it adds an **RBLclass** tab to the Outlook ribbon and a task
pane on the right of the Outlook window.

## Requirements at a glance

- Classic (desktop) Windows Outlook — **not** the "new Outlook" toggle.
  Validated against 32-bit Microsoft 365 Apps for Enterprise on the
  Semi-Annual Enterprise Channel.
- Windows 10 1809 or later.
- No administrator rights — it is a per-user install.

See [Installation](user-guide/installation.md) for the full prerequisites and
step-by-step setup.

!!! note "Internal tool"
    RBLclass is a solo-maintained internal tool. It is distributed to pilot
    users through an internal HTTPS share, not the public web, and is not set up
    for external contributions.
