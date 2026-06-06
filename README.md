# RBLclass v2

RBLclass v2 is the modernised replacement for the legacy VBA-based
RBLclass Outlook macro. It helps users quickly classify emails into
folders across multiple .pst archives, with fast keyword search over
the folder tree, attachment management, and send-time guards.

**Stack**: C# / .NET Framework 4.8 (add-in shell + Outlook adapter)
and .NET Standard 2.0 (business core), packaged as a classic
**Outlook Shared COM Add-in** (`IDTExtensibility2` +
`IRibbonExtensibility`). SQLite with FTS5 for the index. WPF
+ MVVM (CommunityToolkit.Mvvm) for UI hosted in Custom Task Panes.
Per-user install via HKCU registration — no admin rights, no
ClickOnce, no VSTO runtime.

See [CLAUDE.md](CLAUDE.md) for the full architectural rules and
[ROADMAP.md](ROADMAP.md) for the delivery phases.

## Phase 0 — Hello PST POC

`/poc/RBLclass.HelloPstPoc` is a **throwaway** proof-of-concept that
de-risks the three Phase-0 technical concerns before the Phase 1
layered solution is built:

1. A COM add-in we author and sign loads into both 64-bit Outlook
   (this dev machine, Current channel) and 32-bit Outlook (the
   target workstations, Semi-Annual Enterprise channel).
2. `Microsoft.Data.Sqlite` + the native `e_sqlite3.dll` work inside
   the Outlook process at both bitnesses
   (`SQLitePCLRaw.bundle_e_sqlite3` ships both runtimes).
3. A per-user install (PowerShell + HKCU registration, no admin)
   does not trip Stormshield / EDR.

The POC is **not** the Phase 1 codebase. Do not extend it; once it
has done its job it will be deleted in favour of the
`/src/RBLclass.{Core,Outlook.Adapter,AddIn}` solution.

### Prerequisites

- Visual Studio 2022 (Community or higher) with the
  **.NET desktop development** workload and the
  **.NET Framework 4.8 targeting pack**. (The Office/SharePoint
  workload is NOT required.)
- The standalone **.NET SDK** (e.g.
  `winget install Microsoft.DotNet.SDK.8`). Required even though we
  target `net48` — the SDK-style csproj needs `Microsoft.NET.Sdk` to
  resolve. On Windows ARM64 the VS workload does not bundle it.
- **Office must be installed locally** — the build references the
  Office PIAs (`Extensibility`, `office`, `Microsoft.Office.Interop.Outlook`)
  directly from the GAC. They land there when Office is installed.
- At least one PST file attached to your Outlook profile, with a
  non-empty first subfolder.

The deploy script will create a self-signed code-signing cert
automatically the first time it runs.

### Easy-button deploy

Close Outlook. Then from a non-elevated PowerShell anywhere:

```powershell
cd <repo>\poc\scripts
.\Deploy-HelloPstAddIn.ps1
```

This single script runs the full workflow:

1. Finds `msbuild.exe` via `vswhere` (no Developer PowerShell
   needed).
2. Creates a self-signed cert (`CN=RBLclass POC Dev`) if none
   exists, or reuses the existing one.
3. Uninstalls any prior install (idempotent).
4. Clears Outlook **Resiliency** entries that mention our ProgId,
   and adds the add-in to `DoNotDisableAddinList`, so Outlook will
   re-try loading us after a previous crash and not auto-disable
   us next time.
5. Restores + builds the solution (Debug|AnyCPU).
6. Copies the build output to
   `%LocalAppData%\RBLclass\HelloPstPoc\`, Authenticode-signs the
   DLL, and writes the HKCU COM + add-in registry entries.

Start Outlook. Look for the **RBLclass** ribbon tab.

### Step-by-step (if you'd rather not use the wrapper)

```powershell
cd <repo>\poc\scripts
.\Create-SelfSignedCert.ps1                                # prints the thumbprint
cd ..
msbuild RBLclass.HelloPstPoc.sln /restore `
        /p:Configuration=Debug /p:Platform="Any CPU"
cd scripts
.\Install-HelloPstAddIn.ps1 -Thumbprint <thumbprint>
```

After a successful build the output lives under
`poc\RBLclass.HelloPstPoc\bin\Debug\net48\`. Verify both native
SQLite payloads are present:

```
bin\Debug\net48\runtimes\win-x86\native\e_sqlite3.dll
bin\Debug\net48\runtimes\win-x64\native\e_sqlite3.dll
```

### Confirm the install

1. A new ribbon tab **RBLclass** appears with one button **Hello
   PST**.
2. `File → Options → Add-ins` lists "RBLclass Hello PST POC" under
   *Active Application Add-ins*.
3. Clicking **Hello PST** shows a MessageBox with process bitness,
   first PST name, item count in its first subfolder, the loaded
   `e_sqlite3.dll` path, and the SQLite roundtrip row count.

### Deploy to a target workstation (Phase 0 EDR observation)

To package the POC for install on a separate machine (typically
32-bit Outlook on the Semi-Annual Enterprise channel), without
needing the code-signing cert or any dev tooling on that machine:

```powershell
cd poc\scripts
.\Stage-TargetRelease.ps1
```

This rebuilds Release, stages payload + scripts + a tiny on-target
README under `%TEMP%\HelloPstPoc-target\`, and produces
`%USERPROFILE%\Desktop\HelloPstPoc-target.zip` — **AES-256
encrypted with password `rbl-v2`** by default, so Gmail / corporate
mail filters that strip unencrypted zips containing `.ps1` / `.dll`
will let it through. Pass `-Password ''` to skip encryption.

Requires **7-Zip** locally for the encryption step
(`winget install 7zip.7zip`).

Email the zip + the password to whoever runs the target machine.
On the target: extract, close Outlook, and run from the extracted
folder:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\Install-HelloPstAddIn.ps1 -SkipSigning -BuildOutput .\payload
```

Outlook does not require an Authenticode signature to load a COM
add-in — `-SkipSigning` is fine for the Phase 0 observation pass.
The real installer (Phase 1, WiX MSI) will sign with the internal-
PKI cert.

### Uninstall

Close Outlook. Then:

```powershell
cd poc\scripts
.\Uninstall-HelloPstAddIn.ps1
```

What the uninstaller does:

1. Removes the COM-class registration from
   `HKCU:\Software\Classes\CLSID\{guid}` and its `Wow6432Node\`
   mirror.
2. Removes the Outlook add-in entry at
   `HKCU:\Software\Microsoft\Office\Outlook\Addins\RBLclass.HelloPstAddIn`.
3. Deletes `%LocalAppData%\RBLclass\HelloPstPoc\`. Leaves
   `%LocalAppData%\RBLclass\hello-pst-poc.db` in place (it
   contains POC ping rows; remove manually if you want a clean
   slate).

Verify the uninstall:

- Reopen Outlook → no **RBLclass** ribbon tab.
- `File → Options → Add-ins` → "RBLclass Hello PST POC" is gone.
- `HKCU:\Software\Microsoft\Office\Outlook\Addins\` no longer
  contains an `RBLclass.HelloPstAddIn` key.
