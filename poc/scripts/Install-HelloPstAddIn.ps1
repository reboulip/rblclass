<#
.SYNOPSIS
    Install the RBLclass Phase 0 Hello PST COM add-in for the current
    user. No admin rights required.

.DESCRIPTION
    Copies the build output to %LocalAppData%\RBLclass\HelloPstPoc,
    Authenticode-signs the main DLL with the supplied certificate,
    and writes the HKCU registry entries that make Outlook discover
    and load the add-in.

    Outlook must be closed before running. The script refuses to
    proceed if it detects a running outlook.exe.

.PARAMETER Thumbprint
    SHA1 thumbprint of the code-signing certificate in
    Cert:\CurrentUser\My. Use Create-SelfSignedCert.ps1 to make one
    for POC use.

.PARAMETER Configuration
    Build configuration to install from. Default: Debug.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Thumbprint,

    [ValidateSet("Debug", "Release")]
    [string] $Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

# --- Paths --------------------------------------------------------------

$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$pocRoot     = Split-Path -Parent $scriptDir
$projectDir  = Join-Path $pocRoot "RBLclass.HelloPstPoc"
$buildOutput = Join-Path $projectDir "bin\$Configuration\net48"
$installDir  = Join-Path $env:LOCALAPPDATA "RBLclass\HelloPstPoc"
$mainDll     = Join-Path $installDir "RBLclass.HelloPstPoc.dll"

# COM class identity (kept in sync with HelloPstAddIn.cs)
$clsid    = "{B5E7A6F1-3D5A-4A8B-9C2E-1F4E5A8B9D03}"
$progId   = "RBLclass.HelloPstAddIn"
$asmName  = "RBLclass.HelloPstPoc, Version=0.1.0.0, Culture=neutral, PublicKeyToken=null"
$asmClass = "RBLclass.HelloPstPoc.HelloPstAddIn"
$runtime  = "v4.0.30319"

# --- Pre-flight ---------------------------------------------------------

if (-not (Test-Path $buildOutput)) {
    throw "Build output not found at '$buildOutput'. Build the solution first (Debug|AnyCPU) from $pocRoot."
}

if (-not (Test-Path (Join-Path $buildOutput "RBLclass.HelloPstPoc.dll"))) {
    throw "RBLclass.HelloPstPoc.dll missing from '$buildOutput'."
}

$x86Native = Join-Path $buildOutput "runtimes\win-x86\native\e_sqlite3.dll"
$x64Native = Join-Path $buildOutput "runtimes\win-x64\native\e_sqlite3.dll"
if (-not (Test-Path $x86Native)) {
    Write-Warning "Missing $x86Native. 32-bit Outlook will fail at SQLite load."
}
if (-not (Test-Path $x64Native)) {
    Write-Warning "Missing $x64Native. 64-bit Outlook will fail at SQLite load."
}

if (Get-Process -Name outlook -ErrorAction SilentlyContinue) {
    throw "Outlook is running. Close it before installing the add-in."
}

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

# Normalize path to signtool.exe
if ($signtool -is [System.Management.Automation.CommandInfo]) {
    $signtool = $signtool.Source
}
elseif ($signtool -is [System.IO.FileInfo]) {
    $signtool = $signtool.FullName
}


# --- Copy files ---------------------------------------------------------

Write-Host "Installing to $installDir"
if (Test-Path $installDir) {
    Remove-Item $installDir -Recurse -Force
}
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Copy-Item -Path (Join-Path $buildOutput "*") -Destination $installDir -Recurse -Force

# --- Sign the main DLL --------------------------------------------------

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

# --- Helper: write a registry value -------------------------------------

function Set-RegValue {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] $Value,
        [Microsoft.Win32.RegistryValueKind] $Kind = "String"
    )
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    if ($Name -eq "(default)") {
        New-ItemProperty -Path $Path -Name "(default)" -Value $Value -PropertyType $Kind -Force | Out-Null
    } else {
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Kind -Force | Out-Null
    }
}

# --- Register the .NET COM class (HKCU) ---------------------------------
#
# HKCU\Software\Classes is shared between 32-bit and 64-bit clients on
# 64-bit Windows (no WOW64 redirection on HKCU), so we write a single
# subtree. We also write the Wow6432Node mirror defensively in case a
# given Windows/Outlook combo reads it.

$clsidPaths = @(
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid"
)

foreach ($base in $clsidPaths) {
    Write-Host "Writing $base"
    Set-RegValue -Path $base -Name "(default)" -Value $progId

    $inproc = Join-Path $base "InprocServer32"
    Set-RegValue -Path $inproc -Name "(default)"     -Value "mscoree.dll"
    Set-RegValue -Path $inproc -Name "ThreadingModel" -Value "Both"
    Set-RegValue -Path $inproc -Name "Class"          -Value $asmClass
    Set-RegValue -Path $inproc -Name "Assembly"       -Value $asmName
    Set-RegValue -Path $inproc -Name "RuntimeVersion" -Value $runtime
    $codeBaseUri = ([System.Uri](New-Object System.Uri($mainDll))).AbsoluteUri
    Set-RegValue -Path $inproc -Name "CodeBase"       -Value $codeBaseUri

    $verKey = Join-Path $inproc "0.1.0.0"
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
Set-RegValue -Path $progIdPath -Name "(default)" -Value "RBLclass Hello PST POC"
Set-RegValue -Path (Join-Path $progIdPath "CLSID") -Name "(default)" -Value $clsid

# --- Register as an Outlook add-in --------------------------------------

$addinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"
Write-Host "Writing $addinKey"
Set-RegValue -Path $addinKey -Name "FriendlyName"    -Value "RBLclass Hello PST POC"
Set-RegValue -Path $addinKey -Name "Description"     -Value "Phase 0 throwaway POC. See ROADMAP.md."
Set-RegValue -Path $addinKey -Name "LoadBehavior"    -Value 3 -Kind DWord
Set-RegValue -Path $addinKey -Name "CommandLineSafe" -Value 1 -Kind DWord

Write-Host ""
Write-Host "Done. Start Outlook and look for the RBLclass ribbon tab."
