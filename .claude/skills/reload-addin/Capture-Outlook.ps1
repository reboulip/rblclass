# Captures the main Outlook window to a PNG via PrintWindow (works even when
# Outlook is covered by other windows or in the background). Lets Claude
# visually inspect the add-in UI after /reload-addin without asking the user
# to describe or screenshot it. PW_RENDERFULLCONTENT captures DWM-composed
# content, including DirectX/WPF surfaces.
param(
    [string]$OutFile = (Join-Path $env:TEMP "outlook-capture.png")
)

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class RblCapture {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
}
"@

$p = Get-Process outlook -ErrorAction Stop | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if (-not $p) { Write-Error "No Outlook process with a main window."; exit 1 }

$r = New-Object RblCapture+RECT
[RblCapture]::GetWindowRect($p.MainWindowHandle, [ref]$r) | Out-Null
$w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
if ($w -le 0 -or $h -le 0) { Write-Error "Outlook window has an empty rect (minimized?)."; exit 1 }

$bmp = New-Object System.Drawing.Bitmap $w, $h
$g = [System.Drawing.Graphics]::FromImage($bmp)
$hdc = $g.GetHdc()
$ok = [RblCapture]::PrintWindow($p.MainWindowHandle, $hdc, 2)  # 2 = PW_RENDERFULLCONTENT
$g.ReleaseHdc($hdc)
$g.Dispose()
$bmp.Save($OutFile, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Output "printwindow=$ok size=${w}x${h} saved=$OutFile"
