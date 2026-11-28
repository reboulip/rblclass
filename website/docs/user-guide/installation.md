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

!!! warning "Close Outlook before installing"
    Windows cannot replace files Outlook has open. If Outlook is open, Windows
    Installer shows a "close the following applications" window and waits for an
    answer, which can look like a hang if it sits behind another window.
    Choosing the automatic-close option closes Outlook, finishes the install and
    restarts Outlook.

1. Download `RBLclass-<version>.msi` from the internal distribution share (ask
   your deployment team for the link).
2. Double-click the `.msi`. It is a standard per-user Windows Installer package
   — no extraction, no extra steps.
3. A small progress window appears. There is no wizard and nothing to configure
   (a SmartScreen prompt may appear, see below).
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

Close Outlook, then install a newer `RBLclass-<version>.msi` the same way — it
replaces the previous version. The search index under `%LocalAppData%\RBLclass\` is left in place so a
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
   Uninstall**. (Re-running the `.msi` does not offer a **Remove** option —
   there is no wizard.)

The search index database is intentionally **left in place** under
`%LocalAppData%\RBLclass\` so that a reinstall is fast. For a fully clean slate,
delete that folder by hand after uninstalling.

## Troubleshooting the installer

- **Automatic log (version 2.11 and later).** Every install, upgrade and
  uninstall writes a verbose log to `%TEMP%\MSI*.LOG` with no switch needed.
  Open `%TEMP%` and sort by **Date modified**. If no log appears, use the manual
  command below (a machine Logging policy, or a failure before the package
  opens, can prevent it). These logs are not cleaned up automatically.
- **Manual log.** In Command Prompt:

    ```
    msiexec /i "RBLclass-<version>.msi" /l*vx "%TEMP%\rblclass-install.log"
    ```

    In PowerShell use `"$env:TEMP\rblclass-install.log"` for the log path.
    Search the file for `Return value 3` to find the failing action.
- **Event Viewer.** Run `eventvwr.msc`, go to **Windows Logs → Application** and
  filter on Source **MsiInstaller**.
- **Fallback: the PowerShell install kit.** `RBLclass-<version>-target.zip` is
  AES-encrypted; your deployment team supplies the password. With Outlook closed
  (the script refuses to run otherwise), run
  `Install-RblClassAddIn.ps1 -SkipSigning`. To go back to the MSI afterwards, run
  `Uninstall-RblClassAddIn.ps1` first.

The installer log is separate from the runtime log in
`%LocalAppData%\RBLclass\logs\`, which records what the add-in does once Outlook
is running. Its startup line lists the add-in, core and adapter versions. If it
is followed by a warning that the RBLclass assemblies come from different builds,
do a clean reinstall: uninstall, delete only `%LocalAppData%\RBLclass\AddIn`, then
install again.

## Reporting a problem

RBLclass is maintained on [GitHub Issues](https://github.com/reboulip/rblclass/issues).
If you can't reach GitHub from your workstation, use the contact channel from the
installation guide your deployment team provided. To help diagnose a problem,
attach the latest log file from `%LocalAppData%\RBLclass\logs\`.
