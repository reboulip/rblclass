<#
.SYNOPSIS
    Remove the RBLclass product COM add-in from the current user's machine.

.DESCRIPTION
    Reverses Install-RblClassAddIn.ps1 using the same release.config.json:
    removes the Outlook Addins entry, the HKCU COM class registration
    (both views), and the install folder under
    %LocalAppData%\<Product>\<InstallSubdir>.

    The SQLite database under %LocalAppData%\<Product>\ is left in place so
    a reinstall keeps the user's index. Delete it manually for a clean
    slate.

    Outlook must be closed before running.

.PARAMETER ConfigPath
    Explicit path to release.config.json. Default: next to this script.
#>

[CmdletBinding()]
param(
    [string] $ConfigPath
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ConfigPath) { $ConfigPath = Join-Path $scriptDir "release.config.json" }
if (-not (Test-Path $ConfigPath)) {
    throw "release.config.json not found at '$ConfigPath'. Run this script from inside the install kit."
}
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

$clsid      = $config.Clsid
$progId     = $config.ProgId
$installDir = Join-Path $env:LOCALAPPDATA (Join-Path $config.Product $config.InstallSubdir)

if (Get-Process -Name outlook -ErrorAction SilentlyContinue) {
    throw "Outlook is running. Close it before uninstalling the add-in."
}

$paths = @(
    "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid",
    "HKCU:\Software\Classes\$progId"
)

# Also remove the Custom Task Pane host control registration, if present.
if ($config.TaskPaneControl) {
    $ctlClsid = $config.TaskPaneControl.Clsid
    $ctlProg  = $config.TaskPaneControl.ProgId
    $paths += @(
        "HKCU:\Software\Classes\CLSID\$ctlClsid",
        "HKCU:\Software\Classes\Wow6432Node\CLSID\$ctlClsid",
        "HKCU:\Software\Classes\$ctlProg"
    )
}

foreach ($p in $paths) {
    if (Test-Path $p) {
        Write-Host "Removing $p"
        Remove-Item -Path $p -Recurse -Force
    } else {
        Write-Host "Not present: $p"
    }
}

if (Test-Path $installDir) {
    Write-Host "Removing $installDir"
    Remove-Item -Path $installDir -Recurse -Force
} else {
    Write-Host "Not present: $installDir"
}

Write-Host ""
Write-Host "Uninstall complete. Restart Outlook to confirm the ribbon tab is gone."
Write-Host "Note: the SQLite index under %LocalAppData%\$($config.Product)\ is left in place."
