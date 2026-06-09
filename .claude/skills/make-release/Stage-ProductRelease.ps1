<#
.SYNOPSIS
    Build the RBLclass product add-in (NOT the POC) and package it both as
    a PowerShell install kit (.zip) and as the per-user MSI that is the
    artifact actually described in docs/installation-guide.html and
    shipped to pilot users.

.DESCRIPTION
    Product-grade adaptation of the validated POC staging flow
    (poc/scripts/Stage-TargetRelease.ps1). Self-contained and idempotent.
    Runs from any PowerShell (no Developer PowerShell needed):

      0. Reads release.config.json (single source of truth for COM
         identity, paths and payload checks) and hard-fails on the
         all-zeros sentinel Clsid so a half-configured release can never
         ship.
      1. Locates msbuild.exe via vswhere.
      2. Refreshes PATH so a freshly-installed .NET SDK is visible.
      3. Rebuilds the solution (Release|AnyCPU by default).
      4. Stages the add-in build output under <stage>\payload\ and copies
         the Install/Uninstall scripts + release.config.json alongside.
      5. Verifies the main DLL and BOTH x86 and x64 native SQLite payloads
         are present (target is 32-bit Outlook, but the same artifact must
         keep working on 64-bit dev Outlook too).
      6. Zips the staging folder (AES-256 encrypted by default via 7-Zip;
         most corporate mail filters refuse unencrypted zips of .dll/.ps1).
      7. Builds the per-user MSI from the SAME Release output via
         installer\Build-Installer.ps1 (itself driven by
         release.config.json), so the kit and the MSI are always stamped
         with the same version and COM identity and can never drift apart.

    The produced zip is a self-contained install kit for the target
    machine - see the on-target README.txt it writes for the procedure.
    The produced MSI (installer\bin\<Product>-<Version>.msi) is the
    artifact end users actually double-click, per
    docs/installation-guide.html; CLAUDE.md "Packaging" has both coexist
    until the MSI is validated on the 32-bit target.

    This script is invoked by the /make-release skill; it can also be run
    by hand.

.PARAMETER Configuration
    Debug or Release. Default: Release.

.PARAMETER Version
    Optional release version stamped into the kit's README and zip name
    (e.g. 1.0.0). Default: the AssemblyVersion from release.config.json.

.PARAMETER OutputZip
    Path of the .zip to produce.
    Default: %USERPROFILE%\Desktop\RBLclass-<version>-target.zip.

.PARAMETER StageRoot
    Working folder where the unzipped layout is built before zipping.
    Default: %TEMP%\RBLclass-target. Wiped at the start of every run.

.PARAMETER Password
    Password used to AES-256-encrypt the produced zip via 7-Zip. Pass an
    empty string ('') to produce an unencrypted zip via Compress-Archive
    instead. Default: 'rbl-v2'.
#>

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string] $Configuration = "Release",

    [string] $Version,

    [string] $OutputZip,

    [string] $StageRoot = (Join-Path $env:TEMP "RBLclass-target"),

    [AllowEmptyString()]
    [string] $Password = "rbl-v2"
)

$ErrorActionPreference = "Stop"

# ----------------------------------------------------------------------
# Step 0: Load config and gate on the sentinel identity
# ----------------------------------------------------------------------

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$config    = Get-Content (Join-Path $scriptDir "release.config.json") -Raw | ConvertFrom-Json

$sentinelClsid = "{00000000-0000-0000-0000-000000000000}"
if ($config.Clsid -eq $sentinelClsid) {
    throw @"
release.config.json still carries the all-zeros sentinel Clsid.
Before the first real release, reconcile its COM identity with the
add-in class in src\RBLclass.AddIn:
  - Clsid           <- the class's [Guid(...)]
  - ProgId          <- the class's [ProgId(...)]
  - AssemblyName    <- RBLclass.AddIn
  - AssemblyClass   <- fully-qualified class name
  - AssemblyVersion <- the assembly version
Then re-run. This gate exists so a half-configured kit can never ship.
"@
}

if (-not $Version) { $Version = $config.AssemblyVersion }
if (-not $OutputZip) {
    $OutputZip = Join-Path $env:USERPROFILE "Desktop\$($config.Product)-$Version-target.zip"
}

# Locate the repo root from git, falling back to walking up from the skill
# folder (.claude/skills/make-release -> repo root is three levels up).
$repoRoot = & git -C $scriptDir rev-parse --show-toplevel 2>$null
if (-not $repoRoot -or -not (Test-Path $repoRoot)) {
    $repoRoot = (Resolve-Path (Join-Path $scriptDir "..\..\..")).Path
}
$repoRoot = $repoRoot.Trim()

$sln      = Join-Path $repoRoot $config.Solution
$buildOut = Join-Path $repoRoot ($config.AddInProjectOutput -replace "\{Configuration\}", $Configuration)

if (-not (Test-Path $sln)) {
    throw @"
Solution not found at '$sln'.
The Phase 1 product solution does not exist yet (only the Phase 0 POC
under poc/ builds today). /make-release targets the real product; build
the Phase 1 skeleton first, or run the POC's own
poc/scripts/Stage-TargetRelease.ps1 to ship the POC.
"@
}

Write-Host "[0/6] config      : $($config.Product) $Version  ($($config.ProgId))"
Write-Host "      repo root   : $repoRoot"

# ----------------------------------------------------------------------
# Step 1: Locate msbuild via vswhere
# ----------------------------------------------------------------------

# Refresh PATH from registry so a freshly-installed .NET SDK is visible to
# MSBuild's SDK resolver.
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("PATH", "User")

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) {
    throw "vswhere.exe not found at '$vswhere'. Install Visual Studio 2022."
}

$msbuild = & $vswhere -latest -prerelease -products * `
    -requires Microsoft.Component.MSBuild `
    -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
if (-not $msbuild -or -not (Test-Path $msbuild)) {
    throw "MSBuild not found via vswhere."
}
Write-Host "[1/6] msbuild     : $msbuild"

# ----------------------------------------------------------------------
# Step 2: Rebuild (clean output)
# ----------------------------------------------------------------------

Write-Host "[2/6] build       : $sln ($Configuration|Any CPU)"
& $msbuild $sln /restore /t:Rebuild "/p:Configuration=$Configuration" "/p:Platform=Any CPU" /v:minimal /nologo
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE."
}

# ----------------------------------------------------------------------
# Step 3: Stage
# ----------------------------------------------------------------------

Write-Host "[3/6] stage       : $StageRoot"
if (Test-Path $StageRoot) { Remove-Item $StageRoot -Recurse -Force }
$payloadDir = Join-Path $StageRoot "payload"
New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null

if (-not (Test-Path $buildOut)) {
    throw "Build output not found at '$buildOut' after a successful build. Check AddInProjectOutput in release.config.json."
}

Copy-Item -Path (Join-Path $buildOut "*") -Destination $payloadDir -Recurse -Force
Copy-Item -Path (Join-Path $scriptDir "Install-RblClassAddIn.ps1")   -Destination $StageRoot -Force
Copy-Item -Path (Join-Path $scriptDir "Uninstall-RblClassAddIn.ps1") -Destination $StageRoot -Force
Copy-Item -Path (Join-Path $scriptDir "release.config.json")         -Destination $StageRoot -Force

# Drop an on-target README so whoever runs this on the target workstation
# doesn't have to fish out the repo.
$readme = @"
$($config.FriendlyName) - target machine install kit
=====================================================
Version: $Version

1. Close Outlook.
2. In a non-elevated PowerShell in this folder:

     Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
     .\Install-RblClassAddIn.ps1 -SkipSigning

   The installer reads release.config.json from this kit, copies the
   payload\ folder to %LocalAppData%\$($config.Product)\$($config.InstallSubdir),
   and writes the HKCU COM + Outlook add-in registry entries.

3. Start Outlook. The "$($config.FriendlyName)" add-in loads on the
   RBLclass ribbon tab.

To uninstall:

     .\Uninstall-RblClassAddIn.ps1

Signing is OPTIONAL: Outlook does not require an Authenticode signature
to load a COM add-in (validated in Phase 0). To sign with the internal-PKI
code-signing cert, drop -SkipSigning and pass -Thumbprint <sha1>.
"@
Set-Content -Path (Join-Path $StageRoot "README.txt") -Value $readme -Encoding utf8

# ----------------------------------------------------------------------
# Step 4: Sanity-check payloads
# ----------------------------------------------------------------------

Write-Host "[4/6] verify      : main DLL + native SQLite payloads"
if (-not (Test-Path (Join-Path $payloadDir $config.MainDll))) {
    throw "Payload is missing $($config.MainDll) - aborting."
}

$missing = $false
foreach ($rel in $config.NativePayloads) {
    $full = Join-Path $payloadDir $rel
    if (Test-Path $full) {
        Write-Host "  OK      $rel"
    } else {
        Write-Warning "  MISSING $rel - SQLite will fail at that bitness"
        $missing = $true
    }
}
if ($missing) {
    throw "One or more native SQLite payloads are missing. The kit would fail at SQLite load on the target; aborting before zip."
}

# ----------------------------------------------------------------------
# Step 5: Zip
# ----------------------------------------------------------------------

if (Test-Path $OutputZip) { Remove-Item $OutputZip -Force }

if ([string]::IsNullOrEmpty($Password)) {
    Write-Host "[5/6] zip         : $OutputZip (unencrypted)"
    Compress-Archive -Path (Join-Path $StageRoot "*") -DestinationPath $OutputZip -Force
} else {
    # Encrypted zip via 7-Zip. Gmail and most corporate mail filters refuse
    # to deliver unencrypted zips containing .ps1 / .dll, but let
    # AES-encrypted ones through because they can't be scanned.
    $sevenZip = $null
    foreach ($p in @(
        "C:\Program Files\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )) {
        if (Test-Path $p) { $sevenZip = $p; break }
    }
    if (-not $sevenZip) {
        $cmd = Get-Command 7z.exe -ErrorAction SilentlyContinue
        if ($cmd) { $sevenZip = $cmd.Source }
    }
    if (-not $sevenZip) {
        throw "7-Zip not found. Install with 'winget install 7zip.7zip' " +
              "or pass -Password '' to produce an unencrypted zip via " +
              "Compress-Archive instead."
    }

    Write-Host "[5/6] zip         : $OutputZip (AES-256 encrypted, password: $Password)"
    & $sevenZip a -tzip "-p$Password" -mem=AES256 -mx=5 $OutputZip (Join-Path $StageRoot "*") | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "7-Zip failed with exit code $LASTEXITCODE."
    }
}

$zipSize = (Get-Item $OutputZip).Length

# ----------------------------------------------------------------------
# Step 6: Build the MSI from the same Release output
# ----------------------------------------------------------------------
# installer\Build-Installer.ps1 reads the SAME release.config.json, so the
# MSI is stamped with the same AssemblyVersion/COM identity as the kit
# above - they can never drift apart. This is the artifact actually
# described in docs/installation-guide.html and double-clicked by pilots;
# the kit remains the documented fallback until the MSI is also validated
# on the 32-bit target (CLAUDE.md "Packaging").

Write-Host "[6/6] msi         : installer\Build-Installer.ps1 -Configuration $Configuration"
& (Join-Path $repoRoot "installer\Build-Installer.ps1") -Configuration $Configuration | ForEach-Object { Write-Host "        $_" }
$msiPath = Join-Path $repoRoot "installer\bin\$($config.Product)-$($config.AssemblyVersion).msi"

Write-Host ""
Write-Host "Done."
Write-Host "  Product : $($config.FriendlyName) $Version"
Write-Host "  Zip     : $OutputZip ($([Math]::Round($zipSize/1KB)) KB)"
Write-Host "  MSI     : $msiPath"
Write-Host "  Staged  : $StageRoot"
if (-not [string]::IsNullOrEmpty($Password)) {
    Write-Host "  Open with password: $Password"
}
