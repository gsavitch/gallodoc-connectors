param (
    [switch]$ForceKillPort3000
)

$ErrorActionPreference = "Continue"

# 1. Resolve paths
$RepoRoot = Get-Location
$AddinDir = Join-Path $RepoRoot "integrations/word_addin"
$ManifestPath = Join-Path $AddinDir "manifest.xml"

Write-Host "--- Stopping HaloBridge Word Connector ---" -ForegroundColor Cyan

# 2. Stop office-addin-debugging
if (Test-Path $AddinDir) {
    Write-Host "Stopping Word debug session..."
    Push-Location $AddinDir
    try {
        npx office-addin-debugging stop manifest.xml desktop
    } catch {
        Write-Host "No active Office debug session found or failed to stop." -ForegroundColor Gray
    }
    Pop-Location
}

# 3. Optionally kill port 3000
if ($ForceKillPort3000) {
    Write-Host "ForceKillPort3000 passed. Checking for process on port 3000..."
    $Port3000Process = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique
    if ($Port3000Process) {
        Write-Host "Killing process ID $($Port3000Process)..." -ForegroundColor Red
        Stop-Process -Id $Port3000Process -Force
    } else {
        Write-Host "No process found on port 3000."
    }
}

Write-Host "Done." -ForegroundColor Green
