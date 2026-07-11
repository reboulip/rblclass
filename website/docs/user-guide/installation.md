# Installation

RBLclass installs **per-user**, with **no administrator rights**. Everything it
writes lives under your own user profile (`HKCU` in the registry and
`%LocalAppData%\RBLclass\`), so there is no machine-wide footprint and nothing
to approve with an admin account.

## Prerequisites

- **Classic (desktop) Windows Outlook** — *not* the "new Outlook" toggle that
  some Microsoft 365 builds show. RBLclass is validated against 32-bit
  Microsoft 365 Apps for Enterprise on the Semi-Annual Enterprise Channel, and
  also loads into 64-bit Outlook.
- **Windows 10 version 1809** or later.
- Your `.pst` archives connected in Outlook as usual. RBLclass reads whatever
  stores Outlook has open.

!!! warning "New Outlook is not supported"
    If your Outlook has a **New Outlook** toggle in the top-right and it is
    switched on, turn it off. RBLclass is a classic COM add-in and does not load
    into the new Outlook.

## Install

1. Download `RBLclass-<version>.msi` from the internal distribution share (ask
   your deployment team for the link).
2. Double-click the `.msi`. It is a standard per-user Windows Installer package
   — no extraction, no extra steps.
3. Click through **Next → Install → Finish**. There is nothing to configure.
4. Start Outlook and look for the **RBLclass** tab on the ribbon.

### SmartScreen warning

Windows SmartScreen will likely show an *"unrecognised publisher"* warning —
this is expected for an internal pilot tool that is not yet code-signed. Choose
**More info → Run anyway**.

!!! danger "A hard block is different from a warning"
    A SmartScreen *warning* is expected. A hard **block** from your
    antivirus/EDR is **not** expected — report it to your deployment team rather
    than working around it.

## First run

The first time Outlook starts with RBLclass loaded:

- The **RBLclass pane** opens automatically on the right.
- RBLclass indexes your folder tree **in the background** so search is instant.
  On a large multi-PST profile the first index can take a short while; search
  results fill in as it completes. The index is stored locally and reused on
  later starts, so this only happens once.

If you ever create or rename folders directly in Outlook and they don't appear
in search, click **Refresh folders** on the RBLclass ribbon tab to re-scan.

## Where RBLclass keeps its data

Everything is under your user profile:

| What | Location |
| --- | --- |
| Folder search index + settings (SQLite) | `%LocalAppData%\RBLclass\` |
| Log files (rolling) | `%LocalAppData%\RBLclass\logs\` |
| COM registration | `HKCU\Software\Classes\...` and `HKCU\Software\Microsoft\Office\Outlook\Addins\` |

No mail bodies are stored in the index — see the [Security](../security.md) page
for exactly what is and isn't persisted.

## Updating

Install a newer `RBLclass-<version>.msi` the same way — it replaces the previous
version. The search index under `%LocalAppData%\RBLclass\` is left in place so a
reinstall doesn't have to re-index from scratch.

RBLclass can also *tell you* when a newer version exists, if your deployment team
enables it: **Settings → Updates → Check for a newer version at startup**, with a
**Version-check URL** pointing at a plain-text version marker on your internal
share. It is **off by default**, sends no data other than the request for that
marker, and does nothing if the URL is blank. See
[Network posture](../security.md#network-posture) for the details.

## Uninstall

1. Close Outlook.
2. **Settings → Apps → Installed apps → "RBLclass - Email Classifier" →
   Uninstall** (or re-run the `.msi` and choose **Remove**).

The search index database is intentionally **left in place** under
`%LocalAppData%\RBLclass\` so that a reinstall is fast. For a fully clean slate,
delete that folder by hand after uninstalling.

## Reporting a problem

RBLclass is maintained on [GitHub Issues](https://github.com/reboulip/rblclass/issues).
If you can't reach GitHub from your workstation, use the contact channel from the
installation guide your deployment team provided. To help diagnose a problem,
attach the latest log file from `%LocalAppData%\RBLclass\logs\`.
