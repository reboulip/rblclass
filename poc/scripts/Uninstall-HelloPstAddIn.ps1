<#
.SYNOPSIS
    Remove the RBLclass Phase 0 Hello PST COM add-in from the current
    user's machine.

.DESCRIPTION
    Reverses Install-HelloPstAddIn.ps1: removes the Outlook Addins
    entry, the HKCU COM class registration (both views), and the
    install folder under %LocalAppData%\RBLclass\HelloPstPoc.

    The SQLite POC database at %LocalAppData%\RBLclass\hello-pst-poc.db
    is left in place so it can be inspected post-mortem. Delete it
    manually if you want a clean slate.

    Outlook must be closed before running.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$clsid     = "{B5E7A6F1-3D5A-4A8B-9C2E-1F4E5A8B9D03}"
$progId    = "RBLclass.HelloPstAddIn"
$installDir = Join-Path $env:LOCALAPPDATA "RBLclass\HelloPstPoc"

if (Get-Process -Name outlook -ErrorAction SilentlyContinue) {
    throw "Outlook is running. Close it before uninstalling the add-in."
}

$paths = @(
    "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid",
    "HKCU:\Software\Classes\$progId"
)

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
Write-Host "Note: %LocalAppData%\RBLclass\hello-pst-poc.db is left in place."
