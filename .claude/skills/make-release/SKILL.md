---
name: make-release
description: Build the RBLclass product add-in and package it into a shippable, self-contained PowerShell install kit (.zip) for a target workstation. Triggered manually with /make-release. Targets the real Phase 1 product (src/RBLclass.* + RBLclass.sln), NOT the Phase 0 POC under poc/.
---

# /make-release — ship the RBLclass product

Produce a full install kit (`.zip`) that can be copied to a target
workstation and installed without admin rights, for the **actual product**
(not the POC). The kit contains the built add-in payload, the
install/uninstall scripts, and `release.config.json`.

This skill encodes the release/packaging workflow that used to live in
CLAUDE.md ("Build and run" / "Publishing"). The proven mechanics are
adapted from the validated Phase 0 POC staging flow.

## Scope

- **In scope:** the Phase 1 product — `RBLclass.sln`, `src/RBLclass.AddIn`
  (+ `.Core`, `.Outlook.Adapter`). Package format: PowerShell install kit
  (build → stage → verify both native SQLite bitnesses → AES-256 zip).
- **Out of scope:** the throwaway POC under `poc/`. To ship the POC, run
  `poc/scripts/Stage-TargetRelease.ps1` directly instead.
- **Not this skill:** the WiX MSI (Phase 1 Publishing target). The install
  kit is the chosen format here; the MSI is a separate, later deliverable.

## Preconditions (check before running, fail clearly)

1. **Branch.** Work on `develop` or a branch off it (repo rule).
2. **Product exists.** `RBLclass.sln` must exist at the repo root and
   `src/RBLclass.AddIn` must build. If it does not, the product has not
   been built yet (Phase 1 of ROADMAP.md is still open) — **stop** and tell
   the user the kit cannot be produced until the Phase 1 solution exists.
   Do not silently fall back to the POC.
3. **COM identity reconciled.** Open the add-in class in
   `src/RBLclass.AddIn` and confirm its `[Guid(...)]` and `[ProgId(...)]`
   match `Clsid` / `ProgId` in [release.config.json](release.config.json),
   along with `AssemblyName`, `AssemblyClass`, `AssemblyVersion`. The
   config ships with an all-zeros **sentinel** `Clsid`; every script in
   this skill refuses to run until you replace it with the real values.
   This is the one-time reconciliation the first real release requires.
4. **Tooling.** Visual Studio 2022 (MSBuild, located via vswhere) + the
   standalone .NET SDK on PATH. 7-Zip for the encrypted zip (or pass
   `-Password ''` for a plain zip).

## Run

From the repo root, with the user's requested configuration/version:

```powershell
.\.claude\skills\make-release\Stage-ProductRelease.ps1 `
    -Configuration Release `
    -Version 1.0.0
```

Useful parameters (see the script header for the full list):

- `-Configuration Debug|Release` (default `Release`).
- `-Version <x.y.z>` — stamped into the README and zip name; defaults to
  `AssemblyVersion` from the config.
- `-OutputZip <path>` — default `%USERPROFILE%\Desktop\RBLclass-<ver>-target.zip`.
- `-Password <pwd>` — AES-256 zip password (default `rbl-v2`); `''` =
  unencrypted zip via `Compress-Archive`.

What the script does, in order: load + sentinel-gate the config → locate
MSBuild → refresh PATH → `Rebuild` the solution (`Release|AnyCPU`) → stage
the add-in output under `payload\` with the install/uninstall scripts and
config → verify the main DLL **and both** `win-x86` + `win-x64` native
`e_sqlite3.dll` are present (aborts before zipping if any is missing) →
zip.

## After it succeeds

Report to the user:

- the zip path and size, and the password if encrypted;
- the on-target install steps (also written into the kit's `README.txt`):

  ```powershell
  # On the target workstation, Outlook closed, non-elevated PowerShell:
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
  .\Install-RblClassAddIn.ps1 -SkipSigning      # signing is optional
  ```

## Signing (optional)

Authenticode signing is **not** required for Outlook to load a COM add-in
(validated in Phase 0). It is defence-in-depth. To sign on the target,
install the internal-PKI code-signing cert and run the kit's installer
with `-Thumbprint <sha1>` instead of `-SkipSigning`.

## Bitness invariant

The same AnyCPU artifact must load into 32-bit Outlook (target) and 64-bit
Outlook (dev). The staging script's hard check on both native SQLite
payloads enforces this — never ship a kit that failed that check.

## Files in this skill

- [release.config.json](release.config.json) — single source of truth for
  COM identity, paths, and payload checks.
- [Stage-ProductRelease.ps1](Stage-ProductRelease.ps1) — build + stage +
  verify + zip.
- [Install-RblClassAddIn.ps1](Install-RblClassAddIn.ps1) — on-target
  per-user installer (config-driven; no hardcoded identity).
- [Uninstall-RblClassAddIn.ps1](Uninstall-RblClassAddIn.ps1) — reverses the
  install.
