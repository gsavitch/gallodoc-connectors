param(
    [switch]$Live
)

$ErrorActionPreference = "Stop"

$distDir = "integrations/word_addin/dist/word-addin"
$manifestPath = "$distDir/manifest.prod.xml"

Write-Host "`n--- Word Connector Production Smoke Test ---" -ForegroundColor Cyan

# 1. Check local files
Write-Host "[1/2] Verifying local files in $distDir..."
if (-not (Test-Path $manifestPath)) { 
    Write-Error "Missing manifest.prod.xml in $distDir"
    exit 1
}
if (-not (Test-Path "$distDir/taskpane.html")) { 
    Write-Error "Missing taskpane.html in $distDir"
    exit 1
}
if (-not (Test-Path "$distDir/taskpane.bundle.js")) { 
    Write-Error "Missing taskpane.bundle.js in $distDir"
    exit 1
}

# Icons
$icons = @("icon-16.png", "icon-32.png", "icon-64.png", "icon-80.png")
foreach ($icon in $icons) {
    if (-not (Test-Path "$distDir/assets/$icon")) { 
        Write-Error "Missing icon: $icon in $distDir/assets"
        exit 1
    }
}
Write-Host "  Success: All required local files exist." -ForegroundColor Green

# 2. Inspect manifest content
Write-Host "[2/2] Inspecting manifest.prod.xml for production readiness..."
$manifestContent = Get-Content $manifestPath -Raw

if ($manifestContent -match "localhost|127\.0\.0\.1") {
    Write-Error "Security Error: manifest.prod.xml contains references to 'localhost'. It must be production-only."
    exit 1
}

$prodUrl = "https://www.halobridge.ai/word-addin/taskpane.html"
if ($manifestContent -notmatch [regex]::Escape($prodUrl)) {
    Write-Error "Manifest is not pointing to the correct production taskpane URL ($prodUrl)."
    exit 1
}
Write-Host "  Success: Manifest contains no localhost and points to production." -ForegroundColor Green

if ($Live) {
    Write-Host "`n--- Starting Live Connection Checks ---" -ForegroundColor Cyan
    
    $endpoints = @(
        @{ Name = "Taskpane HTML"; Uri = "https://www.halobridge.ai/word-addin/taskpane.html" },
        @{ Name = "Production Manifest"; Uri = "https://www.halobridge.ai/word-addin/manifest.prod.xml" },
        @{ Name = "API Health"; Uri = "https://www.halobridge.ai/api/health/" }
    )

    foreach ($ep in $endpoints) {
        Write-Host "Checking $($ep.Name)..." -NoNewline
        try {
            $resp = Invoke-WebRequest -Uri $ep.Uri -Method Get -UseBasicParsing -TimeoutSec 10
            if ($resp.StatusCode -eq 200) {
                Write-Host " OK." -ForegroundColor Green
                if ($ep.Name -eq "API Health") {
                    $json = $resp.Content | ConvertFrom-Json
                    if ($json.status -ne "ok") {
                        Write-Warning "API Health status is not 'ok': $($json.status)"
                    }
                }
            } else {
                Write-Error "$($ep.Name) returned status $($resp.StatusCode)"
                exit 1
            }
        } catch {
            Write-Error "Failed to reach $($ep.Name): $($_.Exception.Message)"
            exit 1
        }
    }

    Write-Host "Checking Icons..." -NoNewline
    foreach ($icon in $icons) {
        $iconUrl = "https://www.halobridge.ai/word-addin/assets/$icon"
        try {
            $resp = Invoke-WebRequest -Uri $iconUrl -Method Get -UseBasicParsing
            if ($resp.StatusCode -ne 200) {
                Write-Error "Icon unreachable: $iconUrl"
                exit 1
            }
        } catch {
            Write-Error "Failed to reach icon $icon: $($_.Exception.Message)"
            exit 1
        }
    }
    Write-Host " OK." -ForegroundColor Green
    
    Write-Host "`nLive checks COMPLETED SUCCESSFULLY." -ForegroundColor Green
} else {
    Write-Host "`nSkipped live checks. Run with -Live to verify hosted endpoints." -ForegroundColor Yellow
}

Write-Host "`n--- ALL SMOKE TESTS PASSED ---`n" -ForegroundColor Cyan
