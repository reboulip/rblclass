<#
.SYNOPSIS
    Build the RBLclass per-user MSI from installer\Package.wxs.

.DESCRIPTION
    Declarative counterpart to the validated PowerShell install kit
    (.claude\skills\make-release): same per-user HKCU registration, same
    %LocalAppData%\RBLclass\AddIn payload, same COM identity - just
    expressed as Windows Installer components instead of imperative
    registry writes, and packaged as a single double-clickable .msi.

    Reads release.config.json (the documented single source of truth for
    the product's COM identity and build-output layout) so the payload
    directory and version always match what /make-release would stage,
    then invokes the WiX CLI (`wix build`, installed as a per-user dotnet
    global tool: dotnet tool install --global wix).

    Requires the AddIn to have been built first - run
    .\.claude\skills\reload-addin\Reload-Addin.ps1 -NoStart (or a normal
    solution build) for -Configuration Debug, or
    .\.claude\skills\make-release\Stage-ProductRelease.ps1 for Release.

    The PowerShell install kit remains the primary, validated path until
    this MSI has been installed and smoke-tested on the target workstation
    (CLAUDE.md "Packaging").

.PARAMETER Configuration
    Debug or Release. Selects which build-output folder to harvest from.
    Default: Release (the GA packaging target).

.PARAMETER OutputPath
    Where to write the .msi. Default: installer\bin\<Product>-<Version>.msi
#>

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string] $Configuration = "Release",

    [string] $OutputPath
)

$ErrorActionPreference = "Stop"

$installerDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot      = & git -C $installerDir rev-parse --show-toplevel 2>$null
if (-not $repoRoot) { $repoRoot = Split-Path -Parent $installerDir }

$configPath = Join-Path $repoRoot ".claude\skills\make-release\release.config.json"
$config     = Get-Content $configPath -Raw | ConvertFrom-Json

$sentinelClsid = "{00000000-0000-0000-0000-000000000000}"
if ($config.Clsid -eq $sentinelClsid) {
    throw "release.config.json still carries the all-zeros sentinel Clsid - the COM identity baked into installer\Package.wxs would not match the real add-in. Reconcile release.config.json first (see Stage-ProductRelease.ps1's gate)."
}

# --- Locate the build output to harvest ---------------------------------

$payloadDir = Join-Path $repoRoot ($config.AddInProjectOutput -replace '\{Configuration\}', $Configuration)
$mainDll    = Join-Path $payloadDir $config.MainDll

if (-not (Test-Path $mainDll)) {
    throw "Build output not found at '$mainDll'. Build src\RBLclass.AddIn ($Configuration|AnyCPU) first, e.g. .\.claude\skills\reload-addin\Reload-Addin.ps1 -NoStart -Configuration $Configuration"
}

foreach ($rel in $config.NativePayloads) {
    if (-not (Test-Path (Join-Path $payloadDir $rel))) {
        Write-Warning "Missing $rel in '$payloadDir' - the matching Outlook bitness will fail at SQLite load. The MSI will still build, but would ship a broken payload."
    }
}

# --- Locate the wix CLI --------------------------------------------------

$wix = Get-Command wix -ErrorAction SilentlyContinue
if (-not $wix) {
    $dotnetTools = Join-Path $env:USERPROFILE ".dotnet\tools"
    $candidate = Join-Path $dotnetTools "wix.exe"
    if (Test-Path $candidate) { $wix = $candidate }
    else {
        throw "wix CLI not found. Install it as a per-user dotnet global tool: dotnet tool install --global wix (then re-open the shell so PATH picks up $dotnetTools)."
    }
}
if ($wix -is [System.Management.Automation.CommandInfo]) { $wix = $wix.Source }

# --- Build ----------------------------------------------------------------

if (-not $OutputPath) {
    $OutputPath = Join-Path $installerDir "bin\$($config.Product)-$($config.AssemblyVersion).msi"
}
$outDir = Split-Path -Parent $OutputPath
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

Write-Host "Building $OutputPath"
Write-Host "  payload : $payloadDir"
Write-Host "  identity: $($config.Clsid) / $($config.ProgId)"

& $wix build (Join-Path $installerDir "Package.wxs") `
    -d "PayloadDir=$payloadDir" `
    -d "AssemblyVersion=$($config.AssemblyVersion)" `
    -arch x86 `
    -o $OutputPath
if ($LASTEXITCODE -ne 0) {
    throw "wix build failed with exit code $LASTEXITCODE."
}

Write-Host ""
Write-Host "Done. $OutputPath"
Write-Host "Install per-user, no admin rights: msiexec /i `"$OutputPath`""
