<#
.SYNOPSIS
    Build the Hello PST POC and package it into a single .zip that can
    be copied to a target workstation and installed with
    Install-HelloPstAddIn.ps1 -SkipSigning.

.DESCRIPTION
    Self-contained, idempotent. Runs from any PowerShell (no
    Developer PowerShell needed):
      1. Locates msbuild.exe via vswhere
      2. Refreshes PATH so freshly-installed tooling is visible
      3. Builds the solution (Release|AnyCPU by default)
      4. Stages bin\<Configuration>\net48\ under <stage>\payload\
         and copies Install/Uninstall scripts alongside
      5. Verifies both x86 and x64 native SQLite payloads are present
         (the target is 32-bit Outlook, but the same artifact must
         keep working on 64-bit dev Outlook too)
      6. Zips the staging folder

    The produced zip is a self-contained "install kit" for the
    target machine — see ../README.md for the on-target procedure.

.PARAMETER Configuration
    Debug or Release. Default: Release.

.PARAMETER OutputZip
    Path of the .zip to produce. Default:
    %USERPROFILE%\Desktop\HelloPstPoc-target.zip.

.PARAMETER StageRoot
    Working folder where the unzipped layout is built before
    zipping. Default: %TEMP%\HelloPstPoc-target. Wiped at the
    start of every run.

.PARAMETER Password
    Password used to AES-256-encrypt the produced zip via 7-Zip.
    Gmail (and most corporate mail filters) refuses to deliver
    unencrypted zips that contain .ps1/.dll, so the default
    behaviour is to encrypt with the password 'rbl-v2'. Pass an
    empty string ('') to produce an unencrypted zip via
    Compress-Archive instead.
#>

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string] $Configuration = "Release",

    [string] $OutputZip = (Join-Path $env:USERPROFILE "Desktop\HelloPstPoc-target.zip"),

    [string] $StageRoot = (Join-Path $env:TEMP "HelloPstPoc-target"),

    [AllowEmptyString()]
    [string] $Password = "rbl-v2"
)

$ErrorActionPreference = "Stop"

# Refresh PATH from registry so a freshly-installed .NET SDK is
# visible to MSBuild's SDK resolver.
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("PATH", "User")

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pocRoot   = Split-Path -Parent $scriptDir
$sln       = Join-Path $pocRoot "RBLclass.HelloPstPoc.sln"
$buildOut  = Join-Path $pocRoot "RBLclass.HelloPstPoc\bin\$Configuration\net48"

# ----------------------------------------------------------------------
# Step 1: Locate msbuild via vswhere
# ----------------------------------------------------------------------

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
Write-Host "[1/5] msbuild     : $msbuild"

# ----------------------------------------------------------------------
# Step 2: Rebuild (clean output)
# ----------------------------------------------------------------------

Write-Host "[2/5] build       : $sln ($Configuration|Any CPU)"
& $msbuild $sln /restore /t:Rebuild "/p:Configuration=$Configuration" "/p:Platform=Any CPU" /v:minimal /nologo
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE."
}

# ----------------------------------------------------------------------
# Step 3: Stage
# ----------------------------------------------------------------------

Write-Host "[3/5] stage       : $StageRoot"
if (Test-Path $StageRoot) { Remove-Item $StageRoot -Recurse -Force }
$payloadDir = Join-Path $StageRoot "payload"
New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null

Copy-Item -Path (Join-Path $buildOut "*") -Destination $payloadDir -Recurse -Force
Copy-Item -Path (Join-Path $scriptDir "Install-HelloPstAddIn.ps1")   -Destination $StageRoot -Force
Copy-Item -Path (Join-Path $scriptDir "Uninstall-HelloPstAddIn.ps1") -Destination $StageRoot -Force

# Drop a tiny on-target README so whoever runs this on the target
# doesn't have to fish out the repo URL.
$readme = @"
RBLclass Hello PST POC - target machine install kit
====================================================

1. Close Outlook.
2. In a non-elevated PowerShell in this folder:

     Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
     .\Install-HelloPstAddIn.ps1 -SkipSigning -BuildOutput .\payload

3. Start Outlook. New ribbon tab "RBLclass" with a "Hello PST"
   button should appear. Click it: a MessageBox reports Outlook
   bitness, PST enumeration, and a SQLite roundtrip.

To uninstall:

     .\Uninstall-HelloPstAddIn.ps1

This is a Phase 0 throwaway POC, unsigned. Outlook does not require
an Authenticode signature to load a COM add-in.
"@
Set-Content -Path (Join-Path $StageRoot "README.txt") -Value $readme -Encoding utf8

# ----------------------------------------------------------------------
# Step 4: Sanity-check native payloads
# ----------------------------------------------------------------------

Write-Host "[4/5] verify      : native SQLite payloads"
$nativeChecks = @(
    @{ Path = "runtimes\win-x86\native\e_sqlite3.dll"; Note = "32-bit Outlook (target workstations)" },
    @{ Path = "runtimes\win-x64\native\e_sqlite3.dll"; Note = "64-bit Outlook (dev machine)" }
)
$missing = $false
foreach ($c in $nativeChecks) {
    $full = Join-Path $payloadDir $c.Path
    if (Test-Path $full) {
        Write-Host "  OK      $($c.Path)"
    } else {
        Write-Warning "  MISSING $($c.Path) - SQLite will fail on $($c.Note)"
        $missing = $true
    }
}
if (-not (Test-Path (Join-Path $payloadDir "RBLclass.HelloPstPoc.dll"))) {
    throw "Payload is missing RBLclass.HelloPstPoc.dll - aborting."
}

# ----------------------------------------------------------------------
# Step 5: Zip
# ----------------------------------------------------------------------

if (Test-Path $OutputZip) { Remove-Item $OutputZip -Force }

if ([string]::IsNullOrEmpty($Password)) {
    Write-Host "[5/5] zip         : $OutputZip (unencrypted)"
    Compress-Archive -Path (Join-Path $StageRoot "*") -DestinationPath $OutputZip -Force
} else {
    # Encrypted zip via 7-Zip. Gmail and most corporate mail filters
    # refuse to deliver unencrypted zips containing .ps1 / .dll, but
    # let AES-encrypted ones through because they can't be scanned.
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

    Write-Host "[5/5] zip         : $OutputZip (AES-256 encrypted, password: $Password)"
    # -tzip      ZIP format (broadly supported, opens with built-in
    #            Windows / macOS / 7-Zip / WinRAR clients)
    # -mx=5      normal compression
    # -mem=AES256 AES-256 encryption of file contents
    # -p         password (also encrypts filenames inside? no, that'd be -mhe=on;
    #            we leave names readable so the recipient can preview
    #            the install kit's structure before unzipping)
    & $sevenZip a -tzip "-p$Password" -mem=AES256 -mx=5 $OutputZip (Join-Path $StageRoot "*") | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "7-Zip failed with exit code $LASTEXITCODE."
    }
}

$zipSize = (Get-Item $OutputZip).Length
Write-Host ""
Write-Host "Done."
Write-Host "  Zip    : $OutputZip ($([Math]::Round($zipSize/1KB)) KB)"
Write-Host "  Staged : $StageRoot"
if (-not [string]::IsNullOrEmpty($Password)) {
    Write-Host "  Open with password: $Password"
}
if ($missing) {
    Write-Warning "One or more native payloads were missing - see warnings above."
}
