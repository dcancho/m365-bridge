#Requires -RunAsAdministrator
$ErrorActionPreference = "Stop"

$root = Split-Path $PSScriptRoot -Parent
$catalog = Join-Path $root "publish\catalog"
$shareName = "M365Bridge"

if (-not (Test-Path $catalog)) {
    throw "Run scripts\install.ps1 first."
}

$user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
if (-not (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name $shareName -Path $catalog -ReadAccess $user | Out-Null
}

Write-Host "Shared $catalog as \\$env:COMPUTERNAME\$shareName (read access for $user)."
