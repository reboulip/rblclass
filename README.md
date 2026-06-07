# RBLclass v2

RBLclass is an Outlook add-in that helps you file emails into folders
across your `.pst` archives faster: keyword search over your whole
folder tree, one-click classification to one or several folders,
attachment clean-up, and a few send-time reminders (forgotten
attachments, external recipients). It runs entirely on your machine
against your local PST files — no server, no Exchange Online
dependency.

It's a from-scratch C#/.NET rewrite of a legacy VBA Outlook macro of
the same name. **v2.0.0.0** is the first release of the rewrite and is
currently rolling out to a pilot group.

## Features

- **Folder search** — type a few keywords to jump straight to a folder
  anywhere in your PST tree, with collapse/expand and a configurable
  match mode (word-prefix by default, substring as an opt-in).
- **One-click classify** — file the selected mail(s) into one or
  several destination folders at once, with per-destination copies and
  attachment handling on copy.
- **Attachment management** — strip attachments from a filed copy
  without touching the original (or a kept copy, if you keep one).
- **Inline sub-folder creation** — create a destination folder on the
  fly while classifying, with a targeted re-index.
- **Conversation widening** — pull the rest of a conversation into view,
  with a task-completion guard.
- **Send-time guards** — configurable reminders for forgotten
  attachments (keyword list) and external recipients (internal-domain
  allowlist), wired to Outlook's `ItemSend`.
- **Sent-item triage** — a prompt to Class / Delete / Move-to-Inbox /
  Leave a sent item or its whole conversation.
- **Settings** — every option above persists across restarts, backed by
  a local SQLite database under `%LocalAppData%\RBLclass\`.

Once installed, it adds a new **RBLclass** tab to the Outlook ribbon.

## Status

v2.0.0.0 is in **pilot rollout** (Phase 6 of the [roadmap](ROADMAP.md)).
The legacy-parity rebuild (Steps 0–10) is complete and signed off — see
[docs/parity-and-regression.md](docs/parity-and-regression.md) for the
checklist. An open `develop` → `main` PR is gating the `v2.0.0.0` tag on
a 32-bit-target regression pass.

## Requirements

- Classic (desktop) Windows Outlook — **not** the "new Outlook" toggle
  some Microsoft 365 builds show. Validated against 32-bit Microsoft
  365 Apps for Enterprise on the Semi-Annual Enterprise Channel.
- Windows 10 1809 or later.
- No administrator rights — it's a per-user install.

## Installing

Download `RBLclass-<version>.msi` and double-click it — a standard
per-user Windows Installer package, no extraction or extra steps. Click
through **Next → Install → Finish** (there's nothing to configure),
then start Outlook and look for the **RBLclass** tab on the ribbon.

Windows SmartScreen will likely show an "unrecognised publisher"
warning — expected for this internal pilot tool, since it isn't
code-signed yet. Choose **More info → Run anyway**. A hard *block* from
your antivirus/EDR is not expected and should be reported rather than
worked around.

To uninstall: close Outlook, then **Settings → Apps → Installed apps →
"RBLclass - Email Classifier" → Uninstall** (or re-run the `.msi` and
choose **Remove**). The search index database is left in place under
`%LocalAppData%\RBLclass\` so a reinstall doesn't have to re-index from
scratch — delete that folder by hand for a fully clean slate.

For the full walkthrough — first-run indexing, log file locations, and
a caveats table — see the trilingual (EN/DE/FR)
[installation guide](docs/installation-guide.html), which is the same
document sent to pilot users.

## Architecture

A strictly-layered solution, dependencies pointing downward only:

```
RBLclass.AddIn            (.NET FW 4.8)      COM add-in shell — IDTExtensibility2,
                                              IRibbonExtensibility, ribbon, task panes
RBLclass.Outlook.Adapter  (.NET FW 4.8)      COM access to the Outlook object model
RBLclass.Core             (.NET Standard 2.0) business logic — no Outlook/UI/COM deps
```

Packaged as a classic **Outlook Shared COM Add-in** (no VSTO),
registered into Outlook via HKCU. Storage is SQLite (FTS5-capable) in
`%LocalAppData%\RBLclass\`; the UI is WPF with a hand-rolled minimal
MVVM hosted in Custom Task Panes; logging is Serilog with a rolling
file sink.

See [CLAUDE.md](CLAUDE.md) for the full architectural rules (COM
interop and lifetime conventions, threading, bitness, SQLite schema
policy, etc.) and [ROADMAP.md](ROADMAP.md) for the delivery phases and
current status.

## Project layout

```
/src
  /RBLclass.Core              business logic, .NET Standard 2.0
  /RBLclass.Outlook.Adapter   COM adapter, .NET Framework 4.8
  /RBLclass.AddIn             COM add-in shell, .NET Framework 4.8
/installer                    WiX MSI project
/tests
  /RBLclass.Core.Tests        xUnit
/poc                          Phase 0 throwaway POC — not part of the
                              shipped product
/docs                         architecture notes, install guide, roadmap detail
```

## Building and running

Open `RBLclass.sln` in Visual Studio 2022 with the **.NET desktop
development** workload, the **.NET Framework 4.8 targeting pack**, and
the standalone **.NET SDK** installed (the Office/SharePoint workload
is not required — Office itself must be installed locally for the PIA
references). F5 launches Outlook with the add-in attached via *Debug →
Start External Program: outlook.exe*. Full prerequisites and quirks are
in [CLAUDE.md](CLAUDE.md#build-and-run).

Two Claude Code skills cover the day-to-day loop on the dev machine:

- `/reload-addin` — rebuild and reload the add-in into a running
  Outlook to verify a change live.
- `/make-release` — build, package, and verify both the install kit
  (`.zip`) and the per-user MSI shipped to pilots.

## Tests

`RBLclass.Core` is covered by xUnit + FluentAssertions + NSubstitute:

```powershell
dotnet test tests\RBLclass.Core.Tests
```

The Outlook adapter and add-in shell need a live Outlook session and
are validated by hand instead — see the testing strategy in
[ROADMAP.md](ROADMAP.md).

## License

Apache License 2.0 — see [LICENSE](LICENSE).

## Contributing / feedback

This is currently a solo-maintained internal tool, developed
interactively on a single machine — the only place it can be built and
validated against real PST data and the deployment-target Outlook. It
isn't set up for external contributions. Pilot users can report issues
through the contact channel listed in the
[installation guide](docs/installation-guide.html).
