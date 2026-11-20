$ErrorActionPreference = "Stop"

$root = Split-Path $PSScriptRoot -Parent
$addin = Join-Path $root "add-in"
$server = Join-Path $root "server\office-bridge-server"
$output = Join-Path $root "publish"

Get-Process office-bridge-server -ErrorAction SilentlyContinue | Stop-Process -Force

Push-Location $addin
try {
    npm run build
    if ($LASTEXITCODE -ne 0) { throw "Add-in build failed." }
}
finally {
    Pop-Location
}

dotnet publish $server -c Release -r win-x64 -o $output
if ($LASTEXITCODE -ne 0) { throw "Server publish failed." }

$addinOutput = Join-Path $output "addin"
if (Test-Path $addinOutput) { Remove-Item $addinOutput -Recurse -Force }
Copy-Item (Join-Path $addin "dist") $addinOutput -Recurse

Write-Host "Published to $output"
