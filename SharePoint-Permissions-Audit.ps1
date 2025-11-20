<#
.SYNOPSIS
    Comprehensive SharePoint Online Permissions Audit Script

.DESCRIPTION
    Audits permissions across multiple SharePoint sites including site, list, library, folder, and item level permissions.
    Exports results to Excel and/or CSV with modern authentication support.

.PARAMETER SiteUrls
    Array of SharePoint site URLs to audit

.PARAMETER ExportPath
    Path where the audit results will be exported (supports .xlsx and .csv)

.PARAMETER ConfigFile
    Path to JSON configuration file containing site URLs and settings

.PARAMETER IncludeListItems
    Include individual list items in the audit (can be time-consuming)

.PARAMETER IncludeFolders
    Include folders in the audit

.PARAMETER ExpandGroupMembership
    Expand and export SharePoint group membership details

.PARAMETER ExportFormat
    Export format: Excel, CSV, or Both

.PARAMETER LogPath
    Path to log file for detailed logging

.PARAMETER MaxRetries
    Maximum number of retries for failed operations (default: 3)

.EXAMPLE
    .\SharePoint-Permissions-Audit.ps1 -ConfigFile "config.json"

.EXAMPLE
    .\SharePoint-Permissions-Audit.ps1 -SiteUrls @("https://tenant.sharepoint.com/sites/site1") -ExportPath "C:\Temp\Audit.xlsx"

.NOTES
    Author: Michael
    Modified: 2025-11-20
    Version: 2.0
    Requirements:
    - PnP.PowerShell module (2.x or higher)
    - ImportExcel module (for Excel export)
    - Appropriate SharePoint permissions
#>

[CmdletBinding(DefaultParameterSetName = 'Direct')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Direct')]
    [ValidateNotNullOrEmpty()]
    [string[]]$SiteUrls,

    [Parameter(Mandatory = $true, ParameterSetName = 'Direct')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Config')]
    [ValidateNotNullOrEmpty()]
    [string]$ExportPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Config')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ConfigFile,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeListItems = $false,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeFolders = $true,

    [Parameter(Mandatory = $false)]
    [bool]$ExpandGroupMembership = $true,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Excel', 'CSV', 'Both')]
    [string]$ExportFormat = 'Excel',

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3
)

#region Logging Functions

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    # Console output with colors
    switch ($Level) {
        'Info'    { Write-Host $logMessage -ForegroundColor White }
        'Warning' { Write-Warning $Message }
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
        'Success' { Write-Host $logMessage -ForegroundColor Green }
    }

    # File output if log path is specified
    if ($script:LogPath) {
        try {
            Add-Content -Path $script:LogPath -Value $logMessage -ErrorAction SilentlyContinue
        }
        catch {
            Write-Warning "Failed to write to log file: $_"
        }
    }
}

#endregion

#region Module Management

function Install-RequiredModules {
    <#
    .SYNOPSIS
        Installs and imports required PowerShell modules
    #>

    Write-Log "Checking required modules..." -Level Info

    $modules = @(
        @{ Name = 'PnP.PowerShell'; Required = $true }
        @{ Name = 'ImportExcel'; Required = ($ExportFormat -eq 'Excel' -or $ExportFormat -eq 'Both') }
    )

    foreach ($module in $modules) {
        if (-not $module.Required) { continue }

        try {
            if (!(Get-Module -ListAvailable -Name $module.Name)) {
                Write-Log "Installing $($module.Name)..." -Level Warning
                Install-Module -Name $module.Name -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                Write-Log "$($module.Name) installed successfully" -Level Success
            }

            Import-Module $module.Name -ErrorAction Stop
            $moduleInfo = Get-Module $module.Name
            Write-Log "$($module.Name) v$($moduleInfo.Version) loaded" -Level Success
        }
        catch {
            Write-Log "Failed to install/import $($module.Name): $_" -Level Error
            throw
        }
    }
}

#endregion

#region Helper Functions

function Get-HasUniquePermissions {
    <#
    .SYNOPSIS
        Checks if an object has unique permissions (breaks inheritance)
    #>
    param($Object)

    try {
        if ($null -ne $Object -and $Object.HasUniqueRoleAssignments) {
            return $true
        }
    }
    catch {
        return $false
    }
    return $false
}

function Invoke-WithRetry {
    <#
    .SYNOPSIS
        Executes a script block with retry logic for transient failures
    #>
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = $script:MaxRetries,

        [Parameter(Mandatory = $false)]
        [int]$DelaySeconds = 2,

        [Parameter(Mandatory = $false)]
        [string]$OperationName = "Operation"
    )

    $attempt = 1
    $success = $false
    $result = $null

    while (-not $success -and $attempt -le $MaxAttempts) {
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            if ($attempt -lt $MaxAttempts) {
                $waitTime = $DelaySeconds * $attempt
                Write-Log "$OperationName failed (attempt $attempt/$MaxAttempts). Retrying in $waitTime seconds... Error: $_" -Level Warning
                Start-Sleep -Seconds $waitTime
                $attempt++
            }
            else {
                Write-Log "$OperationName failed after $MaxAttempts attempts: $_" -Level Error
                throw
            }
        }
    }

    return $result
}

#endregion

#region Permission Functions

function Get-PermissionDetails {
    <#
    .SYNOPSIS
        Retrieves detailed permission information for a SharePoint object
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Object,

        [Parameter(Mandatory = $true)]
        [string]$ObjectType,

        [Parameter(Mandatory = $true)]
        [string]$ObjectTitle,

        [Parameter(Mandatory = $true)]
        [string]$ObjectUrl,

        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $false)]
        [string]$ParentList = ""
    )

    $permissions = @()

    try {
        # Get role assignments using PnP cmdlet
        $roleAssignments = Get-PnPProperty -ClientObject $Object -Property RoleAssignments

        foreach ($roleAssignment in $roleAssignments) {
            try {
                # Load member and role definition bindings
                $member = Get-PnPProperty -ClientObject $roleAssignment -Property Member
                $roleBindings = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings

                $loginName = $member.LoginName
                $principalType = $member.PrincipalType

                # Determine if it's a user or group
                $memberType = switch ($principalType) {
                    "User" { "User" }
                    "SecurityGroup" { "Security Group" }
                    "SharePointGroup" { "SharePoint Group" }
                    default { $principalType }
                }

                # Get all permission levels
                $permissionLevels = @()
                foreach ($roleDef in $roleBindings) {
                    $permissionLevels += $roleDef.Name
                }

                $permissionString = $permissionLevels -join ", "

                $permissions += [PSCustomObject]@{
                    SiteUrl                = $SiteUrl
                    ParentList             = $ParentList
                    ObjectType             = $ObjectType
                    ObjectTitle            = $ObjectTitle
                    ObjectUrl              = $ObjectUrl
                    HasUniquePermissions   = Get-HasUniquePermissions -Object $Object
                    PrincipalName          = $member.Title
                    PrincipalLoginName     = $loginName
                    PrincipalType          = $memberType
                    PermissionLevels       = $permissionString
                    GrantedThrough         = if ($memberType -like "*Group*") { "Group Membership" } else { "Direct" }
                    AuditDate              = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
            catch {
                Write-Log "Error processing role assignment for $ObjectTitle : $_" -Level Warning
            }
        }
    }
    catch {
        Write-Log "Error getting permissions for $ObjectTitle : $_" -Level Warning
    }

    return $permissions
}

function Get-GroupMembers {
    <#
    .SYNOPSIS
        Retrieves members of a SharePoint group
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string]$GroupName,

        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    $members = @()

    try {
        $users = Invoke-WithRetry -ScriptBlock {
            Get-PnPGroupMember -Identity $GroupId
        } -OperationName "Get members for group $GroupName"

        foreach ($user in $users) {
            $members += [PSCustomObject]@{
                SiteUrl            = $SiteUrl
                GroupName          = $GroupName
                UserName           = $user.Title
                UserEmail          = $user.Email
                UserLoginName      = $user.LoginName
                UserPrincipalType  = $user.PrincipalType
                AuditDate          = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
    catch {
        Write-Log "Error getting members for group $GroupName : $_" -Level Warning
    }

    return $members
}

#endregion

#region Site Audit Functions

function Get-SitePermissions {
    <#
    .SYNOPSIS
        Audits permissions for a single SharePoint site
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $false)]
        [bool]$IncludeListItems = $false,

        [Parameter(Mandatory = $false)]
        [bool]$IncludeFolders = $true
    )

    Write-Log "`n========================================" -Level Info
    Write-Log "Auditing Site: $SiteUrl" -Level Info
    Write-Log "========================================" -Level Info

    $allPermissions = @()
    $groupMembership = @()

    try {
        # Connect to SharePoint site using device login (most reliable for modern auth)
        Write-Log "Connecting to site (browser authentication will open)..." -Level Info

        $connected = Invoke-WithRetry -ScriptBlock {
            Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
            return $true
        } -OperationName "Connect to $SiteUrl"

        if (-not $connected) {
            throw "Failed to connect to $SiteUrl"
        }

        # Verify connection
        $web = Get-PnPWeb
        Write-Log "Connected successfully to: $($web.Title)" -Level Success

        # Site-level permissions
        Write-Log "Processing Site-level permissions..." -Level Info
        $web = Get-PnPWeb -Includes RoleAssignments, HasUniqueRoleAssignments, Title, Url

        $sitePerms = Get-PermissionDetails -Object $web `
            -ObjectType "Site" -ObjectTitle $web.Title -ObjectUrl $SiteUrl -SiteUrl $SiteUrl -ParentList ""
        $allPermissions += $sitePerms

        # Get SharePoint Groups and their members
        if ($ExpandGroupMembership) {
            Write-Log "Processing SharePoint Groups..." -Level Info
            $groups = Get-PnPGroup

            $groupCount = 0
            foreach ($group in $groups) {
                $groupCount++
                Write-Progress -Activity "Processing SharePoint Groups" `
                    -Status "Processing group: $($group.Title)" `
                    -PercentComplete (($groupCount / $groups.Count) * 100)

                Write-Log "  - Processing group: $($group.Title)" -Level Info
                $groupMembership += Get-GroupMembers -GroupId $group.Id -GroupName $group.Title -SiteUrl $SiteUrl
            }
            Write-Progress -Activity "Processing SharePoint Groups" -Completed
        }

        # Get all lists and libraries
        Write-Log "Processing Lists and Libraries..." -Level Info
        $lists = Get-PnPList -Includes RoleAssignments, HasUniqueRoleAssignments, Hidden, BaseType, Title, RootFolder, ItemCount

        $visibleLists = $lists | Where-Object { -not $_.Hidden -or $_.Title -like "User Information List" }
        $listCount = 0

        foreach ($list in $visibleLists) {
            $listCount++
            $percentComplete = [math]::Round(($listCount / $visibleLists.Count) * 100, 2)

            Write-Progress -Activity "Processing Lists and Libraries" `
                -Status "[$listCount/$($visibleLists.Count)] $($list.Title) ($($list.ItemCount) items)" `
                -PercentComplete $percentComplete

            Write-Log "  [$listCount/$($visibleLists.Count)] Auditing: $($list.Title) ($($list.ItemCount) items)" -Level Info

            $listUrl = $SiteUrl.TrimEnd('/') + $list.RootFolder.ServerRelativeUrl

            # List-level permissions (only if unique)
            if ($list.HasUniqueRoleAssignments) {
                $listPerms = Get-PermissionDetails -Object $list `
                    -ObjectType "List/Library" -ObjectTitle $list.Title -ObjectUrl $listUrl -SiteUrl $SiteUrl -ParentList $list.Title
                $allPermissions += $listPerms
            }

            # Get list items (optional - can be time consuming)
            if ($IncludeListItems -or $IncludeFolders) {
                try {
                    Write-Log "    Fetching items..." -Level Info

                    # Use CAML query to get all items with FileRef and FSObjType
                    $items = Invoke-WithRetry -ScriptBlock {
                        Get-PnPListItem -List $list -PageSize 2000 -Fields Id, FileLeafRef, FileRef, FSObjType -ErrorAction Stop
                    } -OperationName "Get items for $($list.Title)"

                    if ($items) {
                        Write-Log "    Found $($items.Count) items, checking for unique permissions..." -Level Info

                        $itemCount = 0
                        $uniqueCount = 0

                        foreach ($item in $items) {
                            $itemCount++

                            if ($itemCount % 100 -eq 0) {
                                Write-Progress -Activity "Processing Lists and Libraries" `
                                    -Status "[$listCount/$($visibleLists.Count)] $($list.Title)" `
                                    -CurrentOperation "Processing items: $itemCount/$($items.Count) ($uniqueCount unique)" `
                                    -PercentComplete $percentComplete
                            }

                            # Check if it's a folder
                            $isFolder = $item.FieldValues["FSObjType"] -eq 1

                            # Skip folders if not included
                            if ($isFolder -and -not $IncludeFolders) {
                                continue
                            }

                            # Skip items if not included
                            if (-not $isFolder -and -not $IncludeListItems) {
                                continue
                            }

                            # Load HasUniqueRoleAssignments property
                            try {
                                $itemObject = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments

                                if ($item.HasUniqueRoleAssignments) {
                                    $uniqueCount++
                                    $itemTitle = if ($item.FieldValues["FileLeafRef"]) { $item.FieldValues["FileLeafRef"] } else { "Item $($item.Id)" }
                                    $itemUrl = $SiteUrl.TrimEnd('/') + $item.FieldValues["FileRef"]
                                    $objectType = if ($isFolder) { "Folder" } else { "Item" }

                                    $itemPerms = Get-PermissionDetails -Object $item `
                                        -ObjectType $objectType -ObjectTitle $itemTitle -ObjectUrl $itemUrl -SiteUrl $SiteUrl -ParentList $list.Title
                                    $allPermissions += $itemPerms
                                }
                            }
                            catch {
                                Write-Log "    Error processing item $itemCount : $_" -Level Warning
                            }
                        }

                        Write-Log "    Completed: $uniqueCount items with unique permissions found" -Level Success
                    }
                }
                catch {
                    Write-Log "Error processing items in $($list.Title): $_" -Level Warning
                }
            }
        }

        Write-Progress -Activity "Processing Lists and Libraries" -Completed

        Write-Log "`nSite audit completed: $SiteUrl" -Level Success
        Write-Log "  - Permission entries found: $($allPermissions.Count)" -Level Info
    }
    catch {
        Write-Log "Error auditing site $SiteUrl : $_" -Level Error
    }
    finally {
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
    }

    return @{
        Permissions      = $allPermissions
        GroupMembership  = $groupMembership
    }
}

#endregion

#region Export Functions

function Export-AuditResults {
    <#
    .SYNOPSIS
        Exports audit results to Excel and/or CSV format
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$Permissions,

        [Parameter(Mandatory = $true)]
        [array]$GroupMembership,

        [Parameter(Mandatory = $true)]
        [string]$ExportPath,

        [Parameter(Mandatory = $true)]
        [string]$ExportFormat,

        [Parameter(Mandatory = $true)]
        [hashtable]$AuditStats
    )

    Write-Log "`n========================================" -Level Info
    Write-Log "Exporting results..." -Level Info
    Write-Log "========================================" -Level Info

    # Prepare export directory
    $exportDir = Split-Path -Path $ExportPath -Parent
    if ($exportDir -and !(Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }

    # Export based on format
    switch ($ExportFormat) {
        'Excel' {
            Export-ToExcel -Permissions $Permissions -GroupMembership $GroupMembership `
                -ExportPath $ExportPath -AuditStats $AuditStats
        }
        'CSV' {
            Export-ToCsv -Permissions $Permissions -GroupMembership $GroupMembership `
                -ExportPath $ExportPath -AuditStats $AuditStats
        }
        'Both' {
            $excelPath = [System.IO.Path]::ChangeExtension($ExportPath, '.xlsx')
            $csvPath = [System.IO.Path]::ChangeExtension($ExportPath, '.csv')

            Export-ToExcel -Permissions $Permissions -GroupMembership $GroupMembership `
                -ExportPath $excelPath -AuditStats $AuditStats
            Export-ToCsv -Permissions $Permissions -GroupMembership $GroupMembership `
                -ExportPath $csvPath -AuditStats $AuditStats
        }
    }
}

function Export-ToExcel {
    param(
        [array]$Permissions,
        [array]$GroupMembership,
        [string]$ExportPath,
        [hashtable]$AuditStats
    )

    try {
        # Ensure .xlsx extension
        if ($ExportPath -notlike "*.xlsx") {
            $ExportPath = [System.IO.Path]::ChangeExtension($ExportPath, '.xlsx')
        }

        # Remove existing file
        if (Test-Path $ExportPath) {
            Remove-Item $ExportPath -Force
        }

        # Create summary
        $summary = @(
            [PSCustomObject]@{ Metric = "Audit Date"; Value = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
            [PSCustomObject]@{ Metric = "Duration"; Value = "$([math]::Round($AuditStats.Duration.TotalMinutes, 2)) minutes" }
            [PSCustomObject]@{ Metric = ""; Value = "" }
            [PSCustomObject]@{ Metric = "Total Sites Audited"; Value = $AuditStats.SiteCount }
            [PSCustomObject]@{ Metric = "Total Permission Entries"; Value = $Permissions.Count }
            [PSCustomObject]@{ Metric = "Unique Permissions Found"; Value = $AuditStats.UniquePermissions }
            [PSCustomObject]@{ Metric = "Direct Permissions"; Value = $AuditStats.DirectPermissions }
            [PSCustomObject]@{ Metric = "Group-Based Permissions"; Value = $AuditStats.GroupPermissions }
            [PSCustomObject]@{ Metric = ""; Value = "" }
            [PSCustomObject]@{ Metric = "Sites with Unique Permissions"; Value = ($Permissions | Where-Object { $_.ObjectType -eq "Site" -and $_.HasUniquePermissions } | Select-Object -ExpandProperty SiteUrl -Unique).Count }
            [PSCustomObject]@{ Metric = "Lists with Unique Permissions"; Value = ($Permissions | Where-Object { $_.ObjectType -eq "List/Library" -and $_.HasUniquePermissions }).Count }
            [PSCustomObject]@{ Metric = "Folders with Unique Permissions"; Value = ($Permissions | Where-Object { $_.ObjectType -eq "Folder" }).Count }
            [PSCustomObject]@{ Metric = "Items with Unique Permissions"; Value = ($Permissions | Where-Object { $_.ObjectType -eq "Item" }).Count }
            [PSCustomObject]@{ Metric = ""; Value = "" }
            [PSCustomObject]@{ Metric = "Total Group Members"; Value = $GroupMembership.Count }
            [PSCustomObject]@{ Metric = "Unique Groups"; Value = ($GroupMembership | Select-Object -ExpandProperty GroupName -Unique).Count }
        )

        # Export Summary
        $summary | Export-Excel -Path $ExportPath -WorksheetName "Summary" `
            -AutoSize -FreezeTopRow -BoldTopRow -TableName "SummaryTable" -MoveToStart

        # Export Permissions
        if ($Permissions.Count -gt 0) {
            Write-Log "Exporting $($Permissions.Count) permission entries to Excel..." -Level Info
            $Permissions | Export-Excel -Path $ExportPath -WorksheetName "Permissions" `
                -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "PermissionsTable"
        }
        else {
            Write-Log "No permissions found to export" -Level Warning
            @([PSCustomObject]@{Message = "No permissions found" }) | Export-Excel -Path $ExportPath -WorksheetName "Permissions"
        }

        # Export Group Membership
        if ($GroupMembership.Count -gt 0) {
            Write-Log "Exporting $($GroupMembership.Count) group memberships to Excel..." -Level Info
            $GroupMembership | Export-Excel -Path $ExportPath -WorksheetName "Group Membership" `
                -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "GroupMembershipTable"
        }

        Write-Log "Results exported to Excel: $ExportPath" -Level Success

        # Open Excel file
        try {
            Invoke-Item $ExportPath
        }
        catch {
            Write-Log "Could not auto-open Excel file: $_" -Level Warning
        }
    }
    catch {
        Write-Log "Error exporting to Excel: $_" -Level Error
        throw
    }
}

function Export-ToCsv {
    param(
        [array]$Permissions,
        [array]$GroupMembership,
        [string]$ExportPath,
        [hashtable]$AuditStats
    )

    try {
        $basePath = [System.IO.Path]::GetFileNameWithoutExtension($ExportPath)
        $exportDir = Split-Path -Path $ExportPath -Parent

        # Export Permissions
        if ($Permissions.Count -gt 0) {
            $permissionsPath = Join-Path $exportDir "$basePath`_Permissions.csv"
            Write-Log "Exporting $($Permissions.Count) permission entries to CSV: $permissionsPath" -Level Info
            $Permissions | Export-Csv -Path $permissionsPath -NoTypeInformation -Encoding UTF8
        }

        # Export Group Membership
        if ($GroupMembership.Count -gt 0) {
            $groupsPath = Join-Path $exportDir "$basePath`_GroupMembership.csv"
            Write-Log "Exporting $($GroupMembership.Count) group memberships to CSV: $groupsPath" -Level Info
            $GroupMembership | Export-Csv -Path $groupsPath -NoTypeInformation -Encoding UTF8
        }

        # Export Summary
        $summaryPath = Join-Path $exportDir "$basePath`_Summary.csv"
        $summary = [PSCustomObject]@{
            AuditDate              = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            DurationMinutes        = [math]::Round($AuditStats.Duration.TotalMinutes, 2)
            SitesAudited           = $AuditStats.SiteCount
            TotalPermissionEntries = $Permissions.Count
            UniquePermissions      = $AuditStats.UniquePermissions
            DirectPermissions      = $AuditStats.DirectPermissions
            GroupPermissions       = $AuditStats.GroupPermissions
            GroupMembers           = $GroupMembership.Count
        }
        $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

        Write-Log "Results exported to CSV files in: $exportDir" -Level Success
    }
    catch {
        Write-Log "Error exporting to CSV: $_" -Level Error
        throw
    }
}

#endregion

#region Main Execution

function Start-SharePointPermissionsAudit {
    <#
    .SYNOPSIS
        Main function to orchestrate the SharePoint permissions audit
    #>

    Write-Log "========================================" -Level Info
    Write-Log "SharePoint Permissions Audit Starting" -Level Info
    Write-Log "========================================" -Level Info

    # Load configuration if specified
    if ($PSCmdlet.ParameterSetName -eq 'Config') {
        Write-Log "Loading configuration from: $ConfigFile" -Level Info
        try {
            $config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
            $script:SiteUrls = $config.SiteUrls

            if ($config.ExportPath) { $script:ExportPath = $config.ExportPath }
            if ($null -ne $config.IncludeListItems) { $script:IncludeListItems = $config.IncludeListItems }
            if ($null -ne $config.IncludeFolders) { $script:IncludeFolders = $config.IncludeFolders }
            if ($null -ne $config.ExpandGroupMembership) { $script:ExpandGroupMembership = $config.ExpandGroupMembership }
            if ($config.ExportFormat) { $script:ExportFormat = $config.ExportFormat }
            if ($config.LogPath) { $script:LogPath = $config.LogPath }

            Write-Log "Configuration loaded: $($SiteUrls.Count) sites to audit" -Level Success
        }
        catch {
            Write-Log "Error loading configuration file: $_" -Level Error
            throw
        }
    }

    # Initialize log file
    if (-not $script:LogPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $script:LogPath = Join-Path $env:TEMP "SharePoint_Audit_$timestamp.log"
    }

    Write-Log "Log file: $script:LogPath" -Level Info

    # Install required modules
    Install-RequiredModules

    # Validate site URLs
    Write-Log "`nValidating $($SiteUrls.Count) site URL(s)..." -Level Info
    foreach ($url in $SiteUrls) {
        if ($url -notmatch '^https://.*\.sharepoint\.com/') {
            Write-Log "Invalid SharePoint URL: $url" -Level Warning
        }
    }

    # Initialize collections
    $allPermissions = @()
    $allGroupMembership = @()

    $startTime = Get-Date

    # Process each site
    $siteCount = 0
    foreach ($siteUrl in $SiteUrls) {
        $siteCount++
        Write-Log "`n### Processing site $siteCount of $($SiteUrls.Count) ###" -Level Info

        Write-Progress -Activity "Auditing SharePoint Sites" `
            -Status "Site $siteCount of $($SiteUrls.Count): $siteUrl" `
            -PercentComplete (($siteCount / $SiteUrls.Count) * 100)

        $result = Get-SitePermissions -SiteUrl $siteUrl `
            -IncludeListItems $IncludeListItems `
            -IncludeFolders $IncludeFolders

        $allPermissions += $result.Permissions
        $allGroupMembership += $result.GroupMembership

        Write-Log "Running total: $($allPermissions.Count) permission entries" -Level Info
    }

    Write-Progress -Activity "Auditing SharePoint Sites" -Completed

    $endTime = Get-Date
    $duration = $endTime - $startTime

    # Calculate statistics
    $uniquePerms = ($allPermissions | Where-Object { $_.HasUniquePermissions -eq $true }).Count
    $directPerms = ($allPermissions | Where-Object { $_.GrantedThrough -eq "Direct" }).Count
    $groupPerms = ($allPermissions | Where-Object { $_.GrantedThrough -eq "Group Membership" }).Count

    $auditStats = @{
        Duration           = $duration
        SiteCount          = $SiteUrls.Count
        UniquePermissions  = $uniquePerms
        DirectPermissions  = $directPerms
        GroupPermissions   = $groupPerms
    }

    # Export results
    if ($allPermissions.Count -gt 0 -or $allGroupMembership.Count -gt 0) {
        Export-AuditResults -Permissions $allPermissions -GroupMembership $allGroupMembership `
            -ExportPath $ExportPath -ExportFormat $ExportFormat -AuditStats $auditStats
    }
    else {
        Write-Log "No data to export" -Level Warning
    }

    # Display summary
    Write-Log "`n========================================" -Level Success
    Write-Log "Audit Completed Successfully!" -Level Success
    Write-Log "========================================" -Level Success
    Write-Log "Summary:" -Level Info
    Write-Log "  - Sites Audited: $($SiteUrls.Count)" -Level Info
    Write-Log "  - Duration: $([math]::Round($duration.TotalMinutes, 2)) minutes" -Level Info
    Write-Log "  - Permission Entries: $($allPermissions.Count)" -Level Info
    Write-Log "  - Unique Permissions: $uniquePerms" -Level Info
    Write-Log "  - Direct Permissions: $directPerms" -Level Info
    Write-Log "  - Group Permissions: $groupPerms" -Level Info
    Write-Log "  - Group Members: $($allGroupMembership.Count)" -Level Info
    Write-Log "  - Log File: $script:LogPath" -Level Info
}

# Execute main function
try {
    Start-SharePointPermissionsAudit
}
catch {
    Write-Log "Fatal error during audit: $_" -Level Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Error
    exit 1
}

#endregion
