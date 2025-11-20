<#
.SYNOPSIS
    One-click launcher for SharePoint Permissions Audit

.DESCRIPTION
    Automatically runs the SharePoint permissions audit using config.json.
    No parameters required - just double-click or run this script.

.EXAMPLE
    .\Start-Audit.ps1

.EXAMPLE
    # Run from Windows Explorer: Right-click -> Run with PowerShell
#>

# Set error action preference
$ErrorActionPreference = "Stop"

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Permissions Audit Tool" -ForegroundColor Cyan
Write-Host "Sports and Spinal Physio" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if config.json exists, if not create it from example
$configPath = Join-Path $scriptDir "config.json"
$exampleConfigPath = Join-Path $scriptDir "config.example.json"

if (-not (Test-Path $configPath)) {
    Write-Host "Config file not found. Creating from example..." -ForegroundColor Yellow

    if (Test-Path $exampleConfigPath) {
        Copy-Item $exampleConfigPath $configPath
        Write-Host "Created config.json from example" -ForegroundColor Green
        Write-Host ""
        Write-Host "⚠️  Please review config.json and update if needed." -ForegroundColor Yellow
        Write-Host "Press any key to continue with current settings..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-Host ""
    }
    else {
        Write-Host "Error: config.example.json not found!" -ForegroundColor Red
        Write-Host "Please ensure all script files are present." -ForegroundColor Red
        pause
        exit 1
    }
}

# Display current configuration
try {
    $config = Get-Content $configPath -Raw | ConvertFrom-Json

    Write-Host "Current Configuration:" -ForegroundColor White
    Write-Host "  Sites to audit: $($config.SiteUrls.Count)" -ForegroundColor Gray
    Write-Host "  Export path: $($config.ExportPath)" -ForegroundColor Gray
    Write-Host "  Export format: $($config.ExportFormat)" -ForegroundColor Gray
    Write-Host "  Include folders: $($config.IncludeFolders)" -ForegroundColor Gray
    Write-Host "  Include items: $($config.IncludeListItems)" -ForegroundColor Gray
    Write-Host ""

    # Check if export directory exists
    $exportDir = Split-Path -Path $config.ExportPath -Parent
    if ($exportDir -and -not (Test-Path $exportDir)) {
        Write-Host "Creating export directory: $exportDir" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
        Write-Host "Export directory created" -ForegroundColor Green
        Write-Host ""
    }
}
catch {
    Write-Host "Error reading configuration: $_" -ForegroundColor Red
    Write-Host "Please check config.json is valid JSON" -ForegroundColor Red
    pause
    exit 1
}

# Confirm before starting
Write-Host "Ready to start audit..." -ForegroundColor Green
Write-Host "This will:" -ForegroundColor White
Write-Host "  1. Connect to $($config.SiteUrls.Count) SharePoint sites" -ForegroundColor Gray
Write-Host "  2. Audit permissions at multiple levels" -ForegroundColor Gray
Write-Host "  3. Export results to Excel" -ForegroundColor Gray
Write-Host ""
Write-Host "Press any key to start or Ctrl+C to cancel..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host ""

# Run the audit
try {
    $auditScript = Join-Path $scriptDir "SharePoint-Permissions-Audit.ps1"

    if (-not (Test-Path $auditScript)) {
        Write-Host "Error: SharePoint-Permissions-Audit.ps1 not found!" -ForegroundColor Red
        pause
        exit 1
    }

    Write-Host "Starting audit..." -ForegroundColor Cyan
    Write-Host ""

    # Execute the audit script
    & $auditScript -ConfigFile $configPath

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Audit completed!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Error during audit execution" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Check the log file for details" -ForegroundColor Yellow
}

# Pause at the end so user can see results
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
