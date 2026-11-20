$ErrorActionPreference = "Stop"

$root = Split-Path $PSScriptRoot -Parent
$publish = Join-Path $root "publish"
$exe = Join-Path $publish "office-bridge-server.exe"
$manifest = Join-Path $publish "addin\manifest.xml"
$catalog = Join-Path $publish "catalog"
$shareName = "M365Bridge"
$shareUrl = "\\$env:COMPUTERNAME\$shareName"
$catalogId = "{4a8432a2-2798-49ee-98f7-70bdb1bc169b}"

if (-not (Test-Path $exe) -or -not (Test-Path $manifest)) {
    throw "Run scripts\publish.ps1 first."
}

New-Item $catalog -ItemType Directory -Force | Out-Null
Copy-Item $manifest (Join-Path $catalog "manifest.xml") -Force

$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
New-Item $catalogKey -Force | Out-Null
Set-ItemProperty $catalogKey -Name "Id" -Value $catalogId
Set-ItemProperty $catalogKey -Name "Url" -Value $shareUrl
Set-ItemProperty $catalogKey -Name "Flags" -Value 1 -Type DWord

$startup = [Environment]::GetFolderPath("Startup")
$shortcutPath = Join-Path $startup "M365 Bridge.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exe
$shortcut.WorkingDirectory = $publish
$shortcut.Save()

if (-not (Get-Process office-bridge-server -ErrorAction SilentlyContinue)) {
    Start-Process explorer.exe -ArgumentList "`"$shortcutPath`""
}

if (Test-Path $shareUrl) {
    Write-Host "Installed. Restart Excel, then add M365 Bridge from Add-ins > My Add-ins > Shared Folder."
}
else {
    Write-Host "Installed, but the catalog folder is not shared yet. From an administrator PowerShell, run:"
    Write-Host "  $(Join-Path $PSScriptRoot 'share-catalog.ps1')"
}
