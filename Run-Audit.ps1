<#
.SYNOPSIS
    Quick launcher for SharePoint Permissions Audit

.DESCRIPTION
    Simple script to run the audit using the configuration file.
    Customize config.json before running.

.EXAMPLE
    .\Run-Audit.ps1

.EXAMPLE
    .\Run-Audit.ps1 -ConfigFile ".\custom-config.json"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigFile = ".\config.json"
)

# Check if config file exists, create from example if missing
if (-not (Test-Path $ConfigFile)) {
    Write-Host "Configuration file not found: $ConfigFile" -ForegroundColor Yellow

    # Try to create from example
    $exampleConfig = ".\config.example.json"
    if (Test-Path $exampleConfig) {
        Write-Host "Creating config.json from example..." -ForegroundColor Yellow
        Copy-Item $exampleConfig $ConfigFile
        Write-Host "Created $ConfigFile successfully" -ForegroundColor Green
        Write-Host "You can edit this file to customize your audit settings" -ForegroundColor Gray
        Write-Host ""
    }
    else {
        Write-Host "Error: Neither $ConfigFile nor config.example.json found!" -ForegroundColor Red
        Write-Host "`nPlease create a config.json file with your settings" -ForegroundColor Yellow
        exit 1
    }
}

# Display configuration
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Permissions Audit" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Configuration: $ConfigFile" -ForegroundColor White

try {
    $config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
    Write-Host "Sites to audit: $($config.SiteUrls.Count)" -ForegroundColor White
    Write-Host "Export format: $($config.ExportFormat)" -ForegroundColor White
    Write-Host "`nStarting audit..." -ForegroundColor Yellow
    Write-Host ""

    # Run the audit script
    & ".\SharePoint-Permissions-Audit.ps1" -ConfigFile $ConfigFile
}
catch {
    Write-Host "Error running audit: $_" -ForegroundColor Red
    exit 1
}
