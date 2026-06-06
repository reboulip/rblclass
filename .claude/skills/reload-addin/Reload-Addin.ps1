<#
.SYNOPSIS
    Local dev loop: rebuild the RBLclass product add-in and reload it into
    Outlook so a freshly-developed change can be verified live.

.DESCRIPTION
    The fast inner loop that mirrors /make-release's mechanics without the
    packaging. Order is deliberate: BUILD FIRST while Outlook is still
    open, so a broken build never closes your Outlook. Only once the build
    succeeds does it close Outlook, install from the build output, and
    relaunch.

      1. Build  src\RBLclass.AddIn (Debug|AnyCPU by default) - Outlook stays open.
      2. Close  Outlook gracefully (force-kill only after a timeout).
      3. Install the build output via the make-release installer
         (Install-RblClassAddIn.ps1 -SkipSigning), reusing the shared
         release.config.json COM identity.
      4. Start  Outlook again (unless -NoStart).

    Targets the REAL product only (RBLclass.sln / src\RBLclass.AddIn). If
    the Phase 1 solution does not exist yet, or release.config.json still
    carries the all-zeros sentinel Clsid, it fails clearly and touches
    nothing.

    Invoked by the /reload-addin skill at the end of a development task so
    the change can be checked in Outlook.

.PARAMETER Configuration
    Debug or Release. Default: Debug (the dev-iteration default).

.PARAMETER Clean
    Use the Rebuild target (clean build) instead of incremental Build.

.PARAMETER NoStart
    Do not relaunch Outlook after installing. Useful if you want to start
    it yourself (e.g. under a debugger).

.PARAMETER CloseTimeoutSeconds
    How long to wait for Outlook to close gracefully before force-killing
    it. Default: 20.
#>

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string] $Configuration = "Debug",

    [switch] $Clean,

    [switch] $NoStart,

    [int] $CloseTimeoutSeconds = 20
)

$ErrorActionPreference = "Stop"

# ----------------------------------------------------------------------
# Step 0: Config + preflight (shared with /make-release)
# ----------------------------------------------------------------------

$scriptDir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$makeReleaseDir = (Resolve-Path (Join-Path $scriptDir "..\make-release")).Path
$configPath     = Join-Path $makeReleaseDir "release.config.json"
$installScript  = Join-Path $makeReleaseDir "Install-RblClassAddIn.ps1"

if (-not (Test-Path $configPath))    { throw "Shared config not found at '$configPath'." }
if (-not (Test-Path $installScript)) { throw "Installer not found at '$installScript'." }

$config = Get-Content $configPath -Raw | ConvertFrom-Json
if ($config.Clsid -eq "{00000000-0000-0000-0000-000000000000}") {
    throw @"
release.config.json still carries the all-zeros sentinel Clsid. Reconcile
its COM identity with the add-in class in src\RBLclass.AddIn before the
add-in can be installed (see the /make-release skill). This gate exists so
a half-configured add-in is never registered into Outlook.
"@
}

# Repo root from git, falling back to walking up from the skill folder
# (.claude/skills/reload-addin -> repo root is three levels up).
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
under poc/ builds today). /reload-addin targets the real product; build
the Phase 1 skeleton first.
"@
}

Write-Host "[0/4] target      : $($config.FriendlyName)  ($Configuration|Any CPU)"

# Resolve outlook.exe up front so we relaunch exactly what was installed.
function Get-OutlookPath {
    foreach ($k in @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
    )) {
        if (Test-Path $k) {
            $p = (Get-ItemProperty -Path $k -Name "(default)" -ErrorAction SilentlyContinue)."(default)"
            if ($p -and (Test-Path $p)) { return $p }
        }
    }
    return "outlook.exe"  # fall back to PATH
}
$outlookExe = Get-OutlookPath

# ----------------------------------------------------------------------
# Step 1: Build (Outlook stays open - a broken build never closes it)
# ----------------------------------------------------------------------

# Refresh PATH so a freshly-installed .NET SDK is visible to MSBuild's
# SDK resolver.
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("PATH", "User")

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) { throw "vswhere.exe not found at '$vswhere'. Install Visual Studio 2022." }
$msbuild = & $vswhere -latest -prerelease -products * `
    -requires Microsoft.Component.MSBuild `
    -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
if (-not $msbuild -or -not (Test-Path $msbuild)) { throw "MSBuild not found via vswhere." }

$target = if ($Clean) { "Rebuild" } else { "Build" }
Write-Host "[1/4] build       : $sln /t:$target"
& $msbuild $sln /restore /t:$target "/p:Configuration=$Configuration" "/p:Platform=Any CPU" /v:minimal /nologo
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE. Outlook left running; nothing changed."
}
if (-not (Test-Path (Join-Path $buildOut $config.MainDll))) {
    throw "$($config.MainDll) not found in '$buildOut' after build. Check AddInProjectOutput in release.config.json."
}

# ----------------------------------------------------------------------
# Step 2: Close Outlook (graceful, then force after the timeout)
# ----------------------------------------------------------------------

$ol = Get-Process -Name outlook -ErrorAction SilentlyContinue
if ($ol) {
    Write-Host "[2/4] close        : Outlook (graceful, up to ${CloseTimeoutSeconds}s)"
    foreach ($p in $ol) { $p.CloseMainWindow() | Out-Null }

    $deadline = (Get-Date).AddSeconds($CloseTimeoutSeconds)
    while ((Get-Process -Name outlook -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
    }

    $still = Get-Process -Name outlook -ErrorAction SilentlyContinue
    if ($still) {
        Write-Warning "Outlook did not close within ${CloseTimeoutSeconds}s (unsaved item dialog?). Forcing."
        $still | Stop-Process -Force
        Start-Sleep -Seconds 2
    }
} else {
    Write-Host "[2/4] close        : Outlook not running"
}

# ----------------------------------------------------------------------
# Step 3: Install from the build output (reuse the make-release installer)
# ----------------------------------------------------------------------

Write-Host "[3/4] install      : $buildOut -> %LocalAppData%\$($config.Product)\$($config.InstallSubdir)"
& $installScript -SkipSigning -PayloadDir $buildOut -ConfigPath $configPath

# ----------------------------------------------------------------------
# Step 4: Relaunch Outlook
# ----------------------------------------------------------------------

if ($NoStart) {
    Write-Host "[4/4] start        : skipped (-NoStart). Launch '$outlookExe' yourself."
} else {
    Write-Host "[4/4] start        : $outlookExe"
    Start-Process -FilePath $outlookExe
}

Write-Host ""
Write-Host "Reloaded $($config.FriendlyName) ($Configuration). Outlook is starting -"
Write-Host "open the RBLclass ribbon tab and verify the change."
