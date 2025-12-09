# Google Apps Script Deploy Script
# Run this to upload src/projection.gs to the spreadsheet

Write-Host "=== Google Apps Script Deploy ===" -ForegroundColor Cyan

# Check if clasp is installed
$claspCheck = npm list -g @google/clasp 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "WARNING - clasp is not installed" -ForegroundColor Yellow
    Write-Host "Please run: .\setup-clasp.ps1 first" -ForegroundColor Yellow
    exit 1
}

# Skip login check - clasp push will show error if not logged in
Write-Host ""
Write-Host "Note: Login will be checked during upload..." -ForegroundColor Cyan

# Check current directory
$currentDir = Get-Location
Write-Host ""
Write-Host "Working directory: $currentDir" -ForegroundColor Cyan

# Check .clasp.json
if (-not (Test-Path ".clasp.json")) {
    Write-Host ""
    Write-Host "ERROR - .clasp.json not found" -ForegroundColor Red
    Write-Host "Please run: .\setup-clasp.ps1 first" -ForegroundColor Yellow
    exit 1
}

# Get script ID
$claspConfig = Get-Content ".clasp.json" | ConvertFrom-Json
Write-Host ""
Write-Host "Script ID: $($claspConfig.scriptId)" -ForegroundColor Cyan

# Check if projection.gs exists
if (-not (Test-Path "src/projection.gs")) {
    Write-Host ""
    Write-Host "ERROR - src/projection.gs not found" -ForegroundColor Red
    exit 1
}

# Deploy
Write-Host ""
Write-Host "Uploading code..." -ForegroundColor Cyan
clasp push

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "SUCCESS - Deploy completed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Check the spreadsheet menu 'mt.innダッシュボード'" -ForegroundColor Cyan
} else {
    Write-Host ""
    Write-Host "ERROR - Deploy failed" -ForegroundColor Red
    Write-Host "Please check the error message above" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "=== Complete ===" -ForegroundColor Cyan
