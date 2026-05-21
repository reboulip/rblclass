# RBLclass Hello PST POC

**Throwaway Phase 0 proof-of-concept.** Do not extend this project.
It exists only to de-risk three concerns before the Phase 1
layered solution is built:

1. A classic Outlook **Shared COM Add-in** we author and sign
   loads into both 64-bit Outlook (dev machine, Current channel)
   and 32-bit Outlook (target workstations, Semi-Annual Enterprise
   channel).
2. `Microsoft.Data.Sqlite` + native `e_sqlite3.dll` work inside
   the Outlook process at both bitnesses (the
   `SQLitePCLRaw.bundle_e_sqlite3` package ships both runtimes).
3. A per-user install (PowerShell + HKCU registration, no admin)
   does not trip Stormshield / EDR.

Once these are validated end-to-end, this folder will be deleted
in favour of the `/src/RBLclass.{Core,Outlook.Adapter,AddIn}`
solution. See [ROADMAP.md](../ROADMAP.md) Phase 1.

The detailed build / install / uninstall instructions live in the
[repo README](../README.md). What follows is POC-specific guidance
that doesn't belong in the root README.

## Identity (stable across builds)

| | Value |
|---|---|
| CLSID    | `{B5E7A6F1-3D5A-4A8B-9C2E-1F4E5A8B9D03}` |
| ProgID   | `RBLclass.HelloPstAddIn` |
| HKCU key | `HKCU\Software\Microsoft\Office\Outlook\Addins\RBLclass.HelloPstAddIn` |
| Install  | `%LocalAppData%\RBLclass\HelloPstPoc\` |
| DB       | `%LocalAppData%\RBLclass\hello-pst-poc.db` |

If you need to change the CLSID, update both
[HelloPstAddIn.cs](RBLclass.HelloPstPoc/HelloPstAddIn.cs) and the
`$clsid` variable in
[scripts/Install-HelloPstAddIn.ps1](scripts/Install-HelloPstAddIn.ps1)
+ [scripts/Uninstall-HelloPstAddIn.ps1](scripts/Uninstall-HelloPstAddIn.ps1)
together.

## What's deliberately not done

- No Serilog, no rolling log file. Errors surface as MessageBoxes
  from a top-level try/catch.
- No WPF, no Custom Task Pane, no MVVM. Output is a single
  MessageBox.
- No production `ComRef<T>` helper. COM objects are released
  manually with `Marshal.ReleaseComObject` in `finally` blocks.
- No FTS5, no migrations, no schema versioning. One `PocPing`
  table with two columns.
- No WiX MSI. The PowerShell installer is the entire packaging
  story for Phase 0.
- No CI build. The POC is built locally and signed locally.

All of the above land in Phase 1 in the real
`RBLclass.{Core,Outlook.Adapter,AddIn}` solution.

## Known limits and quirks

- `runtimes\` from `SQLitePCLRaw.bundle_e_sqlite3` should land in
  the build output automatically with the SDK-style csproj. If
  the post-build directory is missing `runtimes\win-x86\` or
  `runtimes\win-x64\`, copy the missing arch's `e_sqlite3.dll`
  from `%USERPROFILE%\.nuget\packages\sqlitepclraw.lib.e_sqlite3.*\`
  into the build output before running the installer.
- The POC writes COM registration to both
  `HKCU:\Software\Classes\CLSID\{clsid}` and
  `HKCU:\Software\Classes\Wow6432Node\CLSID\{clsid}`. HKCU is
  not subject to WOW64 redirection so the first one is what both
  bitnesses actually use; the Wow6432Node mirror is defensive
  belt-and-braces.
- If Outlook silently fails to load the add-in, check
  `HKCU\Software\Microsoft\Office\Outlook\Addins\RBLclass.HelloPstAddIn\LoadBehavior`.
  Outlook flips it from 3 to 2 if an exception escaped during
  `OnConnection`. Investigate the cause, fix it, set it back to
  3, restart Outlook.
