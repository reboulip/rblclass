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

## Preconditions (check before running, fail clearly)

1. **Branch.** Work on `develop` or a branch off it (repo rule).
2. **Product exists.** `RBLclass.sln` must exist at the repo root and
   `src/RBLclass.AddIn` must build. If it does not, the product has not
   been built yet — **stop** and tell the user the release cannot be
   produced until the solution exists.
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

### Clean up stale packaged artifacts (automatic, every release)

Immediately after the MSI step succeeds, delete every `installer\bin\*.msi`
and `*.wixpdb` that does **not** match the version just built (keep
`install.log`). No confirmation needed — the folder is gitignored, so this
is local-build-output hygiene, not a source change:

```powershell
Get-ChildItem "installer\bin\*" | Where-Object {
    ($_.Extension -in '.msi', '.wixpdb') -and ($_.Name -notmatch [regex]::Escape($Version))
} | Remove-Item -Force
```

## Ship to main, tag, and publish the GitHub release

This step runs **after** the kit and MSI are successfully produced.

### Merge develop into main

Default for this solo repo is CLAUDE.md's direct-merge flow — no PR.

**Sync first.** `main` may have moved on the remote without a local fetch —
e.g. a prior release shipped via a GitHub PR merge instead of this direct
flow. Merging against a stale local `main` produces a divergent merge
commit and a rejected push. Fetch and fast-forward before merging:

```powershell
git fetch origin
git checkout main
git merge --ff-only origin/main   # fails loudly if local main has unique commits — stop and ask if so
```

Then merge and ship:

```powershell
git merge --no-ff develop -m "Ship <version> — <sprint theme>"
git tag -a <version> -m "<version> — <sprint theme>"
git push origin main
git push origin <version>
git checkout develop
```

Only use a PR instead (`gh pr create --base main --head develop`, wait for
merge, *then* tag on the merged `main`) if the user explicitly asks for a
reviewed merge. Don't do both — merging directly and then also opening a PR
for the same content produces a redundant/empty PR.

### Find issues resolved by this release

Scan the `## vX.X.X.X` section in `ROADMAP.md` matching the version. For
every `[#N]` back-reference on a **checked** (`- [x]`) item, sort it into
one of two buckets by the item's trailing text:

- **Pending** — the item carries a verification-pending note (the
  established phrasing is `- shipped, pending tests` or `- pending
  <reporter/target> confirmation`, e.g. "not verifiable on the dev
  machine" or "reporter needs to confirm on their environment"). The fix
  shipped but isn't confirmed working yet.
- **Resolved** — everything else checked. Confirmed working (interactively
  during the sprint, or by the automated test suite).

Unchecked (`- [ ]`) items are not part of either bucket — warn the user the
sprint isn't complete if any remain; their issues are untouched.

### Publish the release

```powershell
gh release create <version> `
  --target main `
  --title "<version> — <sprint theme>" `
  --notes-file <release-notes-path> `
  "<path-to-kit.zip>" `
  "installer\bin\RBLclass-<AssemblyVersion>.msi"
```

Write real release notes (grouped Fixed/Added/Changed, per-item summaries
from the roadmap text) to a scratchpad file first and pass it via
`--notes-file` — don't ship with a placeholder note. Print the release URL.

### Resolve GH issues (automatic, every release)

No confirmation needed — this is a mechanical mapping from roadmap
checkbox state (already reviewed by the user when each item was verified)
to GH issue state:

```powershell
# Resolved bucket: close with a reference to the release
foreach ($n in $resolvedIssues) {
  gh issue comment $n --body "Resolved in [<version>](<release-url>)."
  gh issue close $n
}

# Pending bucket: label + ask for confirmation, leave open
gh label create awaiting-verification --color fbca04 `
  --description "Fix shipped; awaiting reporter or real-target confirmation before closing" 2>$null
foreach ($n in $pendingIssues) {
  gh issue edit $n --add-label awaiting-verification
  gh issue comment $n --body "Fix shipped in <version> — please confirm it resolves this on your end."
}
```

If a PR-based merge was used instead of the direct merge, build the PR's
`Closes #N` list from the **resolved bucket only** — never reference a
pending-bucket issue with `Closes`, since merging would auto-close it
before it's actually confirmed.

### Archive the shipped roadmap section

Shipping is the moment a roadmap section becomes history — `ROADMAP.md` must
only hold **unshipped** work. After the tag and GH release are created:

1. In `ROADMAP.md`, find every `## vX.X.X.X` section whose version is now
   shipped (≤ the released version) and whose items are **all** resolved
   (`- [x]`, or explicitly struck-through/deferred like v2.7.0.0 D1 was).
   Sections with any genuinely open `- [ ]` item stay in `ROADMAP.md` (e.g.
   items marked "shipped, pending tests" are still open). A section with
   every item checked but one or more still carrying a "pending reporter/
   target confirmation" note (see the GH-issue step above — its issue is
   still open with `awaiting-verification`) is a judgment call: archiving
   is fine once the note is gone, but if in doubt, ask before moving it.
2. Move each such section **verbatim** to the end of
   `docs/roadmap-history.md`, appending ` (shipped)` to its `##` header.
   Keep version order in the history file; do not rewrite item text.
3. Update the "Where we are" paragraph in `ROADMAP.md` (date + latest
   shipped wave) so it stays accurate.
4. Commit both files together on `develop` with a message like
   `Archive shipped vX.X.X.X roadmap section to history`.

---

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
