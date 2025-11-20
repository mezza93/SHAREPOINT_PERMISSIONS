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

# Check if config file exists
if (-not (Test-Path $ConfigFile)) {
    Write-Host "Configuration file not found: $ConfigFile" -ForegroundColor Red
    Write-Host "Please create a config.json file based on config.example.json" -ForegroundColor Yellow
    Write-Host "`nExample:" -ForegroundColor Cyan
    Write-Host "  Copy-Item config.example.json config.json" -ForegroundColor White
    Write-Host "  # Edit config.json with your site URLs and settings" -ForegroundColor Gray
    Write-Host "  .\Run-Audit.ps1" -ForegroundColor White
    exit 1
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
