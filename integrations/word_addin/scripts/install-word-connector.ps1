# HaloBridge Word Connector Installer
# This script prepares the environment for local development and testing.

param(
    [switch]$StartDevServer,
    [switch]$Sideload
)

Write-Host "--- HaloBridge Word Connector Setup ---" -ForegroundColor Cyan

$addinDir = Join-Path $PSScriptRoot ".."
Set-Location $addinDir

Write-Host "`n[1/4] Installing dependencies..." -ForegroundColor Yellow
npm install

Write-Host "`n[2/4] Installing Office dev certificates..." -ForegroundColor Yellow
npx office-addin-dev-certs install

Write-Host "`n[3/4] Building the add-in..." -ForegroundColor Yellow
npm run build

Write-Host "`n[4/4] Setup Complete!" -ForegroundColor Green

Write-Host "`nNext Steps:" -ForegroundColor White
Write-Host "1. In one terminal, run: npm start" -ForegroundColor Gray
Write-Host "2. In another terminal, run: npx office-addin-debugging start manifest.xml desktop" -ForegroundColor Gray

if ($StartDevServer) {
    Write-Host "`nLaunching Dev Server..." -ForegroundColor Cyan
    Start-Process powershell -ArgumentList "-NoExit", "-Command", "npm start"
}

if ($Sideload) {
    Write-Host "`nSideloading in Word..." -ForegroundColor Cyan
    npx office-addin-debugging start manifest.xml desktop
}
