$ErrorActionPreference = "Stop"

$shareName = "M365Bridge"
$catalogId = "{4a8432a2-2798-49ee-98f7-70bdb1bc169b}"

Get-Process office-bridge-server -ErrorAction SilentlyContinue | Stop-Process -Force

$shortcut = Join-Path ([Environment]::GetFolderPath("Startup")) "M365 Bridge.lnk"
if (Test-Path $shortcut) { Remove-Item $shortcut }

Remove-Item "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId" -Recurse -ErrorAction SilentlyContinue

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    Remove-SmbShare -Name $shareName -Force -ErrorAction SilentlyContinue
    Write-Host "Uninstalled."
}
else {
    Write-Host "Uninstalled. To remove the shared folder too, run from an administrator PowerShell:"
    Write-Host "  Remove-SmbShare -Name $shareName -Force"
}
