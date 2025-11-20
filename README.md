# SharePoint Online Permissions Audit Tool

A comprehensive PowerShell tool for auditing permissions across multiple SharePoint Online sites, including site-level, list/library, folder, and item-level permissions.

## Features

- **Multi-level Permission Auditing**
  - Site-level permissions
  - List and library permissions
  - Folder permissions
  - Individual item permissions (optional)

- **Group Membership Expansion**
  - Extract all SharePoint group members
  - Track both direct and group-based permissions
  - Identify nested group memberships

- **Flexible Export Options**
  - Excel format with formatted tables
  - CSV format for data analysis
  - Export both formats simultaneously

- **Robust Error Handling**
  - Automatic retry logic for transient failures
  - Detailed logging to file
  - Progress tracking with visual feedback

- **Modern Authentication**
  - Interactive browser-based authentication
  - Support for MFA and conditional access
  - PnP PowerShell modern auth

## Requirements

### PowerShell Modules
- **PnP.PowerShell** (v2.x or higher) - SharePoint Online connectivity
- **ImportExcel** - Excel export functionality (only required for Excel export)

### Permissions
The account running the audit must have:
- Read access to all SharePoint sites being audited
- Ability to read permissions (typically requires Site Collection Administrator or higher)

### System Requirements
- Windows PowerShell 5.1 or PowerShell 7+
- Internet connectivity to SharePoint Online
- Modern web browser (for interactive authentication)

## Installation

### 1. Clone or Download Repository

```powershell
git clone https://github.com/yourusername/sharepoint-permissions-audit.git
cd sharepoint-permissions-audit
```

### 2. Install Required Modules

The script will automatically install required modules, but you can install them manually:

```powershell
Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
Install-Module -Name ImportExcel -Force -AllowClobber -Scope CurrentUser
```

## Quick Start

### Option 1: One-Click Execution (Easiest)

**Just run the script - no configuration needed!**

```powershell
.\Start-Audit.ps1
```

This script:
- ✅ Automatically uses the included `config.json` (pre-configured for Sports & Spinal Physio sites)
- ✅ Creates config from example if missing
- ✅ Shows you the current settings before starting
- ✅ Pauses at the end so you can see results
- ✅ Perfect for double-clicking from Windows Explorer

**For Windows Explorer:** Right-click `Start-Audit.ps1` → **Run with PowerShell**

### Option 2: Using Configuration File

1. **Edit config.json** with your site URLs (already created for you):
   ```powershell
   notepad config.json
   ```

2. **Run the audit:**
   ```powershell
   .\Run-Audit.ps1
   ```

   Or use a custom config:
   ```powershell
   .\Run-Audit.ps1 -ConfigFile ".\my-custom-config.json"
   ```

### Option 3: Direct Command Line (Advanced)

```powershell
.\SharePoint-Permissions-Audit.ps1 `
    -SiteUrls @("https://yourtenant.sharepoint.com/sites/site1") `
    -ExportPath "C:\Temp\Audit.xlsx" `
    -ExportFormat "Excel"
```

## Usage Examples

### Audit Single Site with Excel Export

```powershell
.\SharePoint-Permissions-Audit.ps1 `
    -SiteUrls @("https://contoso.sharepoint.com/sites/TeamSite") `
    -ExportPath "C:\Reports\TeamSite_Permissions.xlsx" `
    -ExportFormat "Excel"
```

### Audit Multiple Sites with CSV Export

```powershell
.\SharePoint-Permissions-Audit.ps1 `
    -SiteUrls @(
        "https://contoso.sharepoint.com/sites/HR",
        "https://contoso.sharepoint.com/sites/Finance",
        "https://contoso.sharepoint.com/sites/IT"
    ) `
    -ExportPath "C:\Reports\MultiSite_Audit" `
    -ExportFormat "CSV"
```

### Include Individual List Items (Comprehensive Audit)

```powershell
.\SharePoint-Permissions-Audit.ps1 `
    -SiteUrls @("https://contoso.sharepoint.com/sites/Sensitive") `
    -ExportPath "C:\Reports\Detailed_Audit.xlsx" `
    -IncludeListItems $true `
    -IncludeFolders $true `
    -ExportFormat "Both"
```

### Use Configuration File

```powershell
.\SharePoint-Permissions-Audit.ps1 -ConfigFile ".\my-config.json"
```

## Parameters

### Required Parameters (Direct Mode)

| Parameter | Type | Description |
|-----------|------|-------------|
| `-SiteUrls` | String[] | Array of SharePoint site URLs to audit |
| `-ExportPath` | String | Path where results will be exported |

### Required Parameters (Config Mode)

| Parameter | Type | Description |
|-----------|------|-------------|
| `-ConfigFile` | String | Path to JSON configuration file |

### Optional Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-IncludeListItems` | Boolean | `$false` | Include individual list items (can be slow) |
| `-IncludeFolders` | Boolean | `$true` | Include folders with unique permissions |
| `-ExpandGroupMembership` | Boolean | `$true` | Export SharePoint group membership details |
| `-ExportFormat` | String | `Excel` | Export format: `Excel`, `CSV`, or `Both` |
| `-LogPath` | String | Auto-generated | Path to detailed log file |
| `-MaxRetries` | Integer | `3` | Number of retry attempts for failed operations |

## Output Files

### Excel Export (Default)

The Excel file contains multiple worksheets:

1. **Summary** - High-level audit statistics
   - Audit date and duration
   - Total sites, permissions, and groups
   - Breakdown by permission type

2. **Permissions** - Detailed permission entries
   - Site URL and object details
   - Principal (user/group) information
   - Permission levels
   - Direct vs. group-based permissions

3. **Group Membership** - SharePoint group members
   - Group name and site
   - User details (name, email, login)
   - Principal type

### CSV Export

When using CSV format, three files are created:

- `{filename}_Summary.csv` - Audit statistics
- `{filename}_Permissions.csv` - Permission details
- `{filename}_GroupMembership.csv` - Group membership

## Configuration File Format

```json
{
  "SiteUrls": [
    "https://tenant.sharepoint.com/sites/site1",
    "https://tenant.sharepoint.com/sites/site2"
  ],
  "ExportPath": "C:\\Temp\\Audit.xlsx",
  "IncludeListItems": false,
  "IncludeFolders": true,
  "ExpandGroupMembership": true,
  "ExportFormat": "Excel",
  "LogPath": "C:\\Temp\\audit.log",
  "MaxRetries": 3
}
```

## Performance Considerations

### Audit Duration

Typical audit times:
- **Site-level only**: 30 seconds - 2 minutes per site
- **With folders**: 2-10 minutes per site (depends on content volume)
- **With list items**: 10-60+ minutes per site (can be very slow for large lists)

### Recommendations

1. **Start Small**: Test with a single site first
2. **Avoid Item-Level Initially**: Set `IncludeListItems` to `$false` unless absolutely needed
3. **Schedule Large Audits**: For comprehensive audits, run during off-hours
4. **Use Folders Only**: Most permission issues occur at folder level, not individual items

## Troubleshooting

### Authentication Issues

**Problem**: Browser authentication window doesn't appear

**Solution**:
```powershell
# Clear PnP connection cache
Disconnect-PnPOnline
Clear-PnPConnection
```

### Module Installation Errors

**Problem**: Cannot install modules due to permissions

**Solution**:
```powershell
# Install for current user only
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

### Large List Timeouts

**Problem**: Timeout errors on lists with many items

**Solution**:
- Set `IncludeListItems` to `$false`
- Increase `MaxRetries` parameter
- Audit problematic sites separately

### Memory Issues

**Problem**: Script runs out of memory on very large tenants

**Solution**:
- Audit sites in smaller batches
- Export to CSV instead of Excel
- Increase available memory or run on a more powerful machine

## Best Practices

1. **Document Your Audits**
   - Keep configuration files for recurring audits
   - Store results with date-stamped folders
   - Maintain audit log history

2. **Regular Auditing Schedule**
   - Monthly audits for compliance
   - Post-migration audits
   - After organizational changes

3. **Review and Remediate**
   - Identify over-permissioned users
   - Remove unnecessary unique permissions
   - Standardize permission levels

4. **Security Considerations**
   - Protect audit results (contain sensitive info)
   - Use read-only service accounts when possible
   - Log all audit activities

## Output Analysis Tips

### Finding Over-Permissioned Users

```powershell
# Import permissions from Excel/CSV
$permissions = Import-Csv "Audit_Permissions.csv"

# Find users with Full Control
$fullControlUsers = $permissions | Where-Object {
    $_.PermissionLevels -like "*Full Control*"
}

# Count permissions per user
$permissions | Group-Object PrincipalLoginName |
    Sort-Object Count -Descending |
    Select-Object Count, Name
```

### Identify Broken Inheritance

```powershell
# Items with unique permissions
$uniquePermissions = $permissions | Where-Object {
    $_.HasUniquePermissions -eq $true
}

# Group by site
$uniquePermissions | Group-Object SiteUrl |
    Select-Object Count, Name
```

## Version History

### Version 2.0 (2025-11-20)
- Added configuration file support
- Improved error handling with retry logic
- Added progress tracking
- Added CSV export option
- Enhanced logging functionality
- Performance improvements
- Better parameter validation

### Version 1.0 (Initial)
- Basic permission auditing
- Excel export
- Group membership expansion

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Test thoroughly
4. Submit a pull request

## License

This project is provided as-is for use within your organization. Modify as needed.

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Review log files for detailed error messages
3. Open an issue on GitHub with:
   - PowerShell version
   - Module versions
   - Error messages
   - Steps to reproduce

## Acknowledgments

- Built with [PnP.PowerShell](https://pnp.github.io/powershell/)
- Uses [ImportExcel](https://github.com/dfinke/ImportExcel) for Excel export

## Security Notice

⚠️ **Important**: Audit results contain sensitive information about your SharePoint permissions structure.

- Store results securely
- Limit access to audit files
- Follow your organization's data handling policies
- Consider encrypting audit results
- Do not share audit results outside your organization

## Author

**Michael**
Sports and Spinal Physio
November 2025

---

**Last Updated**: 2025-11-20
