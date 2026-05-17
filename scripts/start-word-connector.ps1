param (
    [switch]$ForceKillPort3000,
    [switch]$SkipInstall,
    [switch]$SkipBuild,
    [switch]$NoSideload,
    [string]$BaseUrl = "https://www.halobridge.ai"
)

$ErrorActionPreference = "Stop"

# 1. Resolve paths
$RepoRoot = Get-Location
$AddinDir = Join-Path $RepoRoot "integrations/word_addin"
$ManifestPath = Join-Path $AddinDir "manifest.xml"
$PackageJson = Join-Path $AddinDir "package.json"

Write-Host "--- HaloBridge Word Connector Launcher ---" -ForegroundColor Cyan

# 2. Validate prerequisites
Write-Host "[1/10] Validating prerequisites..."
if (-not (Test-Path $AddinDir)) { throw "Word add-in directory not found at $AddinDir" }
if (-not (Test-Path $ManifestPath)) { throw "manifest.xml not found at $ManifestPath" }
if (-not (Test-Path $PackageJson)) { throw "package.json not found at $PackageJson" }

try {
    npm --version | Out-Null
} catch {
    throw "npm is not installed or not in PATH."
}

# 3. Install/update dependencies
if (-not $SkipInstall) {
    Write-Host "[2/10] Installing dependencies..." -ForegroundColor Gray
    Push-Location $AddinDir
    npm install
    Pop-Location
} else {
    Write-Host "[2/10] Skipping npm install."
}

# 4. Install/trust Office dev certs
Write-Host "[3/10] Ensuring Office dev certs are trusted..."
Push-Location $AddinDir
npx office-addin-dev-certs install --staticRuntime
Pop-Location

# 5. Build the add-in
if (-not $SkipBuild) {
    Write-Host "[4/10] Building add-in..." -ForegroundColor Gray
    Push-Location $AddinDir
    npm run build
    Pop-Location
} else {
    Write-Host "[4/10] Skipping build."
}

# 6. Stop any existing Office add-in debug session
Write-Host "[5/10] Stopping any existing debug sessions..."
Push-Location $AddinDir
try {
    npx office-addin-debugging stop manifest.xml desktop
} catch {
    # Ignore errors if no session is running
}
Pop-Location

# 7. Handle port 3000
Write-Host "[6/10] Checking port 3000..."
$Port3000Process = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique
if ($Port3000Process) {
    $ProcessDetail = Get-Process -Id $Port3000Process
    Write-Host "Port 3000 is currently in use by process ID $($Port3000Process) ($($ProcessDetail.ProcessName))." -ForegroundColor Yellow
    if ($ForceKillPort3000) {
        Write-Host "ForceKillPort3000 passed. Killing process..." -ForegroundColor Red
        Stop-Process -Id $Port3000Process -Force
    } else {
        Write-Host "Please close the process or use -ForceKillPort3000 to automate this." -ForegroundColor Red
        return
    }
}

# 8. Start webpack dev server
Write-Host "[7/10] Starting dev server in a new window..."
$StartServerCmd = "cd '$AddinDir'; npm start"
Start-Process powershell -ArgumentList "-NoExit", "-Command", "$StartServerCmd"

# 9. Wait for dev server
Write-Host "[8/10] Waiting for dev server at https://localhost:3000/taskpane.html..." -ForegroundColor Gray
$TimeoutSec = 30
$StartTime = Get-Date
$Ready = $false

# Disable SSL certificate validation for polling
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

while (((Get-Date) - $StartTime).TotalSeconds -lt $TimeoutSec) {
    try {
        $Response = Invoke-WebRequest -Uri "https://localhost:3000/taskpane.html" -UseBasicParsing -TimeoutSec 2
        if ($Response.StatusCode -eq 200) {
            $Ready = $true
            break
        }
    } catch {
        # Keep waiting
    }
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 2
}
Write-Host ""

if (-not $Ready) {
    Write-Host "Timeout: Dev server did not become ready at https://localhost:3000/taskpane.html within $TimeoutSec seconds." -ForegroundColor Red
    Write-Host "If this is a certificate issue, ensure 'npx office-addin-dev-certs install' succeeded and you trusted the CA." -ForegroundColor Yellow
    return
}

Write-Host "Dev server is up!" -ForegroundColor Green

# 10. Sideload Word add-in
if (-not $NoSideload) {
    Write-Host "[9/10] Sideloading add-in into Word..."
    Push-Location $AddinDir
    npx office-addin-debugging start manifest.xml desktop
    Pop-Location
} else {
    Write-Host "[9/10] Skipping sideload."
}

# 11. Print final instructions
Write-Host ""
Write-Host "--- SETUP COMPLETE ---" -ForegroundColor Cyan
Write-Host "1. Word should open automatically."
Write-Host "2. Ribbon: Look for 'HaloBridge' tab or 'Open GalloDoc' button."
Write-Host "3. Task Pane: Verify Base URL is: $BaseUrl"
Write-Host "4. Click 'Test' then 'Connect'."
Write-Host "5. Change Mode to 'Free Connected' or 'Enterprise Connected'."
Write-Host "6. Click 'Create Local GalloDoc' or 'Save to HaloBridge'."
Write-Host ""
Write-Host "Running at: https://localhost:3000"
Write-Host "To stop: .\scripts\stop-word-connector.ps1" -ForegroundColor Gray
