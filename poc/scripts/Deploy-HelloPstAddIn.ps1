<#
.SYNOPSIS
    End-to-end build + install of the RBLclass Phase 0 Hello PST POC.

.DESCRIPTION
    Self-contained, idempotent workflow that:
      1. Locates msbuild.exe via vswhere (no Developer PowerShell needed)
      2. Creates a self-signed code-signing cert if -CertSubject doesn't
         already exist in Cert:\CurrentUser\My
      3. Uninstalls any prior install (so the next install starts clean)
      4. Clears Outlook Resiliency entries that mention our ProgId,
         so Outlook re-tries loading the add-in after a previous crash
      5. Restores + builds the solution (Debug|AnyCPU by default)
      6. Installs the add-in (per-user HKCU registration, signed DLL)

    Runs from any PowerShell. No admin rights required.

.PARAMETER Configuration
    Debug or Release. Default: Debug.

.PARAMETER OutlookVersion
    Outlook registry version (16.0 covers Office 2016, 2019, 2021,
    365). Default: 16.0.

.PARAMETER CertSubject
    Subject of the self-signed code-signing cert. Default:
    "CN=RBLclass POC Dev". The script reuses an existing cert with
    this subject if one is found.
#>

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string] $Configuration = "Debug",

    [string] $OutlookVersion = "16.0",

    [string] $CertSubject = "CN=RBLclass POC Dev"
)

$ErrorActionPreference = "Stop"

# Refresh PATH from the registry so freshly-installed tooling (e.g. the
# .NET SDK at C:\Program Files\dotnet) is visible to MSBuild's SDK
# resolver even when this shell was launched before the install.
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("PATH", "User")

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pocRoot   = Split-Path -Parent $scriptDir
$sln       = Join-Path $pocRoot "RBLclass.HelloPstPoc.sln"
$progId    = "RBLclass.HelloPstAddIn"

# ----------------------------------------------------------------------
# Pre-flight: Outlook must be closed
# ----------------------------------------------------------------------

if (Get-Process -Name outlook -ErrorAction SilentlyContinue) {
    throw "Outlook is running. Close it before deploying."
}

# ----------------------------------------------------------------------
# Step 1: Locate msbuild via vswhere
# ----------------------------------------------------------------------

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) {
    throw "vswhere.exe not found at '$vswhere'. Install Visual Studio 2022 (any edition)."
}

$msbuild = & $vswhere -latest -prerelease -products * `
    -requires Microsoft.Component.MSBuild `
    -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
if (-not $msbuild -or -not (Test-Path $msbuild)) {
    throw "MSBuild not found via vswhere. Install the 'MSBuild' VS component."
}
Write-Host "[1/6] msbuild     : $msbuild"

# ----------------------------------------------------------------------
# Step 2: Ensure a code-signing certificate exists
# ----------------------------------------------------------------------

function Get-PocCert {
    Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -eq $CertSubject -and $_.HasPrivateKey } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1
}

$cert = Get-PocCert
if (-not $cert) {
    Write-Host "[2/6] cert        : creating $CertSubject (none found)"
    & (Join-Path $scriptDir "Create-SelfSignedCert.ps1") -Subject $CertSubject | Out-Null
    $cert = Get-PocCert
    if (-not $cert) {
        throw "Could not create or locate the code-signing certificate."
    }
} else {
    $exp = $cert.NotAfter.ToString('yyyy-MM-dd')
    Write-Host "[2/6] cert        : reusing $($cert.Thumbprint) (expires $exp)"
}

# ----------------------------------------------------------------------
# Step 3: Uninstall prior install (idempotent)
# ----------------------------------------------------------------------

Write-Host "[3/6] uninstall   : removing prior install if present"
& (Join-Path $scriptDir "Uninstall-HelloPstAddIn.ps1") | Out-Host

# ----------------------------------------------------------------------
# Step 4: Clear Outlook Resiliency entries that mention our ProgId
# ----------------------------------------------------------------------

Write-Host "[4/6] resiliency  : clearing entries for $progId"

$base = "HKCU:\Software\Microsoft\Office\$OutlookVersion\Outlook\Resiliency"
$progIdBytes = [System.Text.Encoding]::Unicode.GetBytes($progId)

function Test-BytesContain {
    param(
        [byte[]] $Haystack,
        [byte[]] $Needle
    )
    if ($Needle.Length -eq 0 -or $Haystack.Length -lt $Needle.Length) {
        return $false
    }
    $max = $Haystack.Length - $Needle.Length
    for ($i = 0; $i -le $max; $i++) {
        $match = $true
        for ($j = 0; $j -lt $Needle.Length; $j++) {
            if ($Haystack[$i + $j] -ne $Needle[$j]) {
                $match = $false
                break
            }
        }
        if ($match) { return $true }
    }
    return $false
}

foreach ($leaf in @("DisabledItems", "CrashingAddinList")) {
    $path = Join-Path $base $leaf
    if (-not (Test-Path $path)) { continue }
    $regKey = Get-Item $path
    foreach ($name in $regKey.GetValueNames()) {
        $data = $regKey.GetValue($name)
        if ($data -is [byte[]] -and (Test-BytesContain -Haystack $data -Needle $progIdBytes)) {
            Remove-ItemProperty -Path $path -Name $name -Force
            Write-Host "  removed $leaf\$name"
        }
    }
}

# Whitelist our add-in so Outlook will not auto-disable it on crash.
# Useful during dev iteration. Phase 1 production install should NOT
# do this — disabling is a real safety mechanism.
$doNotDisable = Join-Path $base "DoNotDisableAddinList"
if (-not (Test-Path $doNotDisable)) {
    New-Item -Path $doNotDisable -Force | Out-Null
}
Set-ItemProperty -Path $doNotDisable -Name $progId -Value 1 -Type DWord
Write-Host "  whitelisted $progId in DoNotDisableAddinList"

# ----------------------------------------------------------------------
# Step 5: Build
# ----------------------------------------------------------------------

Write-Host "[5/6] build       : $sln ($Configuration|Any CPU)"
& $msbuild $sln /restore "/p:Configuration=$Configuration" "/p:Platform=Any CPU" /v:minimal /nologo
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE."
}

# ----------------------------------------------------------------------
# Step 6: Install
# ----------------------------------------------------------------------

Write-Host "[6/6] install     : copying files, signing, writing HKCU"
& (Join-Path $scriptDir "Install-HelloPstAddIn.ps1") `
    -Thumbprint $cert.Thumbprint `
    -Configuration $Configuration | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "Install failed with exit code $LASTEXITCODE."
}

Write-Host ""
Write-Host "OK. Start Outlook and look for the RBLclass ribbon tab."
