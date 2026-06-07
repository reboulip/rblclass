---
name: make-release
description: Build the RBLclass product add-in and package it as both a self-contained PowerShell install kit (.zip) and the per-user MSI that pilots actually double-click (per docs/installation-guide.html). Triggered manually with /make-release. Targets the real Phase 1 product (src/RBLclass.* + RBLclass.sln), NOT the Phase 0 POC under poc/.
---

# /make-release — ship the RBLclass product

Produce, from one command, **both** shippable artifacts for the **actual
product** (not the POC):

- an install **kit** (`.zip`) that can be copied to a target workstation
  and installed without admin rights — the originally-validated path and
  documented fallback; and
- the per-user **MSI** (`installer\bin\RBLclass-<version>.msi`) — the
  artifact end users actually receive and double-click, as described in
  the trilingual pilot install guide
  (`docs/installation-guide.html`).

Both are built from the *same* Release output and the *same*
`release.config.json` identity, so they can never carry different
versions or COM identities. CLAUDE.md ("Packaging") has them coexist
until the MSI is also validated by install/uninstall on the real 32-bit
target — at that point the kit can retire.

This skill encodes the release/packaging workflow that used to live in
CLAUDE.md ("Build and run" / "Publishing"). The proven mechanics are
adapted from the validated Phase 0 POC staging flow; the MSI build reuses
the declarative WiX packaging added alongside it
(`installer/Package.wxs` + `installer/Build-Installer.ps1`).

## Scope

- **In scope:** the Phase 1 product — `RBLclass.sln`, `src/RBLclass.AddIn`
  (+ `.Core`, `.Outlook.Adapter`). Package formats produced in one run:
  PowerShell install kit (build → stage → verify both native SQLite
  bitnesses → AES-256 zip) **and** the WiX MSI (same Release build →
  `wix build` against `installer/Package.wxs`).
- **Out of scope:** the throwaway POC under `poc/`. To ship the POC, run
  `poc/scripts/Stage-TargetRelease.ps1` directly instead.

## Preconditions (check before running, fail clearly)

1. **Branch.** Work on `develop` or a branch off it (repo rule).
2. **Product exists.** `RBLclass.sln` must exist at the repo root and
   `src/RBLclass.AddIn` must build. If it does not, the product has not
   been built yet — **stop** and tell the user the release cannot be
   produced until the Phase 1 solution exists. Do not silently fall back
   to the POC.
3. **COM identity reconciled.** Open the add-in class in
   `src/RBLclass.AddIn` and confirm its `[Guid(...)]` and `[ProgId(...)]`
   match `Clsid` / `ProgId` in [release.config.json](release.config.json),
   along with `AssemblyName`, `AssemblyClass`, `AssemblyVersion`. The
   config ships with an all-zeros **sentinel** `Clsid`; every script in
   this skill (and `installer\Build-Installer.ps1`) refuses to run until
   you replace it with the real values. This is the one-time
   reconciliation the first real release requires.
4. **Tooling.** Visual Studio 2022 (MSBuild, located via vswhere) + the
   standalone .NET SDK on PATH. 7-Zip for the encrypted zip (or pass
   `-Password ''` for a plain zip). The **`wix` CLI** for the MSI step —
   a per-user dotnet global tool: `dotnet tool install --global wix`
   (then re-open the shell so `%USERPROFILE%\.dotnet\tools` is on PATH).

## Run

From the repo root, with the user's requested configuration/version:

```powershell
.\.claude\skills\make-release\Stage-ProductRelease.ps1 `
    -Configuration Release `
    -Version 2.0.0
```

Useful parameters (see the script header for the full list):

- `-Configuration Debug|Release` (default `Release`) — also selects which
  build output the MSI step harvests, so kit and MSI always match.
- `-Version <x.y.z>` — stamped into the kit's README and zip name;
  defaults to `AssemblyVersion` from the config. The MSI's own filename is
  always derived from `AssemblyVersion` directly (matches the strong-name
  baked into its COM registration — see CLAUDE.md's COM-identity notes).
- `-OutputZip <path>` — default `%USERPROFILE%\Desktop\RBLclass-<ver>-target.zip`.
- `-Password <pwd>` — AES-256 zip password (default `rbl-v2`); `''` =
  unencrypted zip via `Compress-Archive`.

What the script does, in order: load + sentinel-gate the config → locate
MSBuild → refresh PATH → `Rebuild` the solution (`Release|AnyCPU`) → stage
the add-in output under `payload\` with the install/uninstall scripts and
config → verify the main DLL **and both** `win-x86` + `win-x64` native
`e_sqlite3.dll` are present (aborts before zipping if any is missing) →
zip → **build the MSI** via `installer\Build-Installer.ps1
-Configuration <Configuration>` from that same verified Release output.

## After it succeeds

Report to the user, for **both** artifacts:

- **Kit** — the zip path and size, and the password if encrypted; the
  on-target install steps (also written into the kit's `README.txt`):

  ```powershell
  # On the target workstation, Outlook closed, non-elevated PowerShell:
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
  .\Install-RblClassAddIn.ps1 -SkipSigning      # signing is optional
  ```

- **MSI** — the path (`installer\bin\RBLclass-<AssemblyVersion>.msi`) and
  size; the on-target install step (also what
  `docs/installation-guide.html` walks pilots through): double-click the
  `.msi`, accept the SmartScreen "more info → Run anyway" prompt if shown
  (the file is unsigned per CLAUDE.md "Signing is optional"), no admin
  rights required. `msiexec /i "<path>"` is the scriptable equivalent.

If a stale MSI for a previous version exists in `installer\bin\`, point it
out — it's easy to grab the wrong artifact by accident; clean it up if the
user confirms it's no longer needed (the folder is gitignored, so this is
just local-build-output hygiene, not a source change).

## Signing (optional)

Authenticode signing is **not** required for Outlook to load a COM add-in
(validated in Phase 0) — for *either* artifact. It is defence-in-depth.
To sign the kit's installer on the target, install the internal-PKI
code-signing cert and run it with `-Thumbprint <sha1>` instead of
`-SkipSigning`. The MSI build does not currently sign the `.msi` itself —
that remains a manual `signtool` step against the internal-PKI cert if/when
the user wants it (see CLAUDE.md "Packaging").

## Bitness invariant

The same AnyCPU artifact must load into 32-bit Outlook (target) and 64-bit
Outlook (dev), and `Build-Installer.ps1` packages it `-arch x86` (the MSI
architecture flag governs installer behaviour, not the payload's
bitness — the payload is the same AnyCPU DLL either way). The staging
script's hard check on both native SQLite payloads — which gates the MSI
step too, since it harvests the same verified output — enforces this:
never ship a kit or MSI that failed that check.

## Files in this skill

- [release.config.json](release.config.json) — single source of truth for
  COM identity, paths, and payload checks. Read by both this skill's
  scripts AND `installer\Build-Installer.ps1` / `installer\Package.wxs`.
- [Stage-ProductRelease.ps1](Stage-ProductRelease.ps1) — build + stage +
  verify + zip + (as of the relaunch) build the MSI.
- [Install-RblClassAddIn.ps1](Install-RblClassAddIn.ps1) — on-target
  per-user installer for the kit (config-driven; no hardcoded identity).
- [Uninstall-RblClassAddIn.ps1](Uninstall-RblClassAddIn.ps1) — reverses the
  kit install.
- `..\..\..\installer\Build-Installer.ps1` / `Package.wxs` (outside this
  skill folder, alongside the WiX project) — the declarative MSI build
  this skill now drives as its final step. Can also be run standalone
  (e.g. to rebuild just the MSI after a version bump without re-zipping a
  kit) — see its own header for `-Configuration` / `-OutputPath`.
