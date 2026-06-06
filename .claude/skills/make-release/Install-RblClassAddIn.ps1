<#
.SYNOPSIS
    Install the RBLclass product COM add-in for the current user. No admin
    rights required.

.DESCRIPTION
    Product installer (NOT the POC). Reads the COM identity, paths and
    payload location from release.config.json sitting next to this script
    inside the install kit, copies the payload to
    %LocalAppData%\<Product>\<InstallSubdir>, optionally Authenticode-signs
    the main DLL, and writes the HKCU registry entries that make Outlook
    discover and load the add-in.

    Outlook must be closed before running. The script refuses to proceed
    if it detects a running outlook.exe.

.PARAMETER Thumbprint
    SHA1 thumbprint of the code-signing certificate in Cert:\CurrentUser\My.
    Required unless -SkipSigning is passed.

.PARAMETER SkipSigning
    Skip Authenticode signing. The add-in still loads (no signature is
    required to load a COM add-in - validated in Phase 0).

.PARAMETER PayloadDir
    Explicit path to the build-output payload folder containing the main
    DLL + dependencies + runtimes\. Default: the .\payload folder produced
    by Stage-ProductRelease.ps1 next to this script.

.PARAMETER ConfigPath
    Explicit path to release.config.json. Default: next to this script.
#>

[CmdletBinding(DefaultParameterSetName = 'Sign')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Sign')]
    [string] $Thumbprint,

    [Parameter(Mandatory = $true, ParameterSetName = 'NoSign')]
    [switch] $SkipSigning,

    [string] $PayloadDir,

    [string] $ConfigPath
)

$ErrorActionPreference = "Stop"

# --- Config -------------------------------------------------------------

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ConfigPath) { $ConfigPath = Join-Path $scriptDir "release.config.json" }
if (-not (Test-Path $ConfigPath)) {
    throw "release.config.json not found at '$ConfigPath'. Run this script from inside the install kit produced by Stage-ProductRelease.ps1."
}
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

if ($config.Clsid -eq "{00000000-0000-0000-0000-000000000000}") {
    throw "release.config.json carries the all-zeros sentinel Clsid; this kit was built without a real COM identity and must not be installed."
}

if (-not $PayloadDir) { $PayloadDir = Join-Path $scriptDir "payload" }

# --- COM class identity (from config) -----------------------------------

$clsid    = $config.Clsid
$progId   = $config.ProgId
$asmName  = "$($config.AssemblyName), Version=$($config.AssemblyVersion), Culture=neutral, PublicKeyToken=null"
$asmClass = $config.AssemblyClass
$asmVer   = $config.AssemblyVersion
$runtime  = "v4.0.30319"

$installDir = Join-Path $env:LOCALAPPDATA (Join-Path $config.Product $config.InstallSubdir)
$mainDll    = Join-Path $installDir $config.MainDll

# --- Pre-flight ---------------------------------------------------------

if (-not (Test-Path $PayloadDir)) {
    throw "Payload folder not found at '$PayloadDir'."
}
if (-not (Test-Path (Join-Path $PayloadDir $config.MainDll))) {
    throw "$($config.MainDll) missing from '$PayloadDir'."
}

foreach ($rel in $config.NativePayloads) {
    if (-not (Test-Path (Join-Path $PayloadDir $rel))) {
        Write-Warning "Missing $rel in payload. The matching Outlook bitness will fail at SQLite load."
    }
}

if (Get-Process -Name outlook -ErrorAction SilentlyContinue) {
    throw "Outlook is running. Close it before installing the add-in."
}

if (-not $SkipSigning) {
    $cert = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Thumbprint -eq $Thumbprint } |
        Select-Object -First 1
    if (-not $cert) {
        throw "Certificate with thumbprint '$Thumbprint' not found in Cert:\CurrentUser\My."
    }

    $signtool = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if (-not $signtool) {
        $sdkRoot = "${env:ProgramFiles(x86)}\Windows Kits\10\bin"
        if (Test-Path $sdkRoot) {
            $candidate = Get-ChildItem $sdkRoot -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -match "x64" } |
                Select-Object -First 1
            if ($candidate) { $signtool = $candidate }
        }
        if (-not $signtool) {
            throw "signtool.exe not found. Install the Windows SDK or add it to PATH."
        }
    }
    if ($signtool -is [System.Management.Automation.CommandInfo]) { $signtool = $signtool.Source }
    elseif ($signtool -is [System.IO.FileInfo])                  { $signtool = $signtool.FullName }
}

# --- Copy files ---------------------------------------------------------

Write-Host "Installing to $installDir"
if (Test-Path $installDir) { Remove-Item $installDir -Recurse -Force }
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Copy-Item -Path (Join-Path $PayloadDir "*") -Destination $installDir -Recurse -Force

# --- Sign the main DLL --------------------------------------------------

if ($SkipSigning) {
    Write-Host "Skipping Authenticode signing (-SkipSigning)"
} else {
    Write-Host "Signing $mainDll with thumbprint $Thumbprint"
    & $signtool sign `
        /sha1 $Thumbprint `
        /fd sha256 `
        /tr "http://timestamp.digicert.com" `
        /td sha256 `
        $mainDll
    if ($LASTEXITCODE -ne 0) {
        throw "signtool sign failed with exit code $LASTEXITCODE."
    }
}

# --- Helper: write a registry value -------------------------------------

function Set-RegValue {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] $Value,
        [Microsoft.Win32.RegistryValueKind] $Kind = "String"
    )
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Kind -Force | Out-Null
}

# --- Register the .NET COM class (HKCU) ---------------------------------
#
# HKCU\Software\Classes is shared between 32-bit and 64-bit clients on
# 64-bit Windows (no WOW64 redirection on HKCU), so a single subtree
# suffices. We also write the Wow6432Node mirror defensively in case a
# given Windows/Outlook combo reads it.

$clsidPaths = @(
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid"
)

foreach ($base in $clsidPaths) {
    Write-Host "Writing $base"
    Set-RegValue -Path $base -Name "(default)" -Value $progId

    $inproc = Join-Path $base "InprocServer32"
    Set-RegValue -Path $inproc -Name "(default)"      -Value "mscoree.dll"
    Set-RegValue -Path $inproc -Name "ThreadingModel" -Value "Both"
    Set-RegValue -Path $inproc -Name "Class"          -Value $asmClass
    Set-RegValue -Path $inproc -Name "Assembly"       -Value $asmName
    Set-RegValue -Path $inproc -Name "RuntimeVersion" -Value $runtime
    $codeBaseUri = ([System.Uri](New-Object System.Uri($mainDll))).AbsoluteUri
    Set-RegValue -Path $inproc -Name "CodeBase"       -Value $codeBaseUri

    $verKey = Join-Path $inproc $asmVer
    Set-RegValue -Path $verKey -Name "Class"          -Value $asmClass
    Set-RegValue -Path $verKey -Name "Assembly"       -Value $asmName
    Set-RegValue -Path $verKey -Name "RuntimeVersion" -Value $runtime
    Set-RegValue -Path $verKey -Name "CodeBase"       -Value $codeBaseUri

    Set-RegValue -Path (Join-Path $base "ProgId") -Name "(default)" -Value $progId

    $impl = Join-Path $base "Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
    if (-not (Test-Path $impl)) { New-Item -Path $impl -Force | Out-Null }
}

# ProgId -> CLSID mapping
$progIdPath = "HKCU:\Software\Classes\$progId"
Set-RegValue -Path $progIdPath -Name "(default)" -Value $config.FriendlyName
Set-RegValue -Path (Join-Path $progIdPath "CLSID") -Name "(default)" -Value $clsid

# --- Register as an Outlook add-in --------------------------------------

$addinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"
Write-Host "Writing $addinKey"
Set-RegValue -Path $addinKey -Name "FriendlyName"    -Value $config.FriendlyName
Set-RegValue -Path $addinKey -Name "Description"     -Value $config.Description
Set-RegValue -Path $addinKey -Name "LoadBehavior"    -Value 3 -Kind DWord
Set-RegValue -Path $addinKey -Name "CommandLineSafe" -Value 1 -Kind DWord

Write-Host ""
Write-Host "Done. Start Outlook and look for the RBLclass ribbon tab."
