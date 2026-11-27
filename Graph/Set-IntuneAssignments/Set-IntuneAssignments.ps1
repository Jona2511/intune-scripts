#Requires -Version 5.1

<#
.SYNOPSIS
    Bulk-assigns Intune apps and policies to groups, from a CSV or parameters.

.DESCRIPTION
    Supported object types (-Type / CSV column Type):
      App               deviceAppManagement/mobileApps
      DeviceConfiguration  deviceManagement/deviceConfigurations   (templates, custom OMA-URI)
      Compliance        deviceManagement/deviceCompliancePolicies
      SettingsCatalog   deviceManagement/configurationPolicies     (incl. newer Endpoint security)
      AdminTemplate     deviceManagement/groupPolicyConfigurations
      PlatformScript    deviceManagement/deviceManagementScripts

    The Intune /assign actions REPLACE all assignments of an object. This script therefore
    reads the existing assignments first, merges the new ones in and posts the complete set,
    so existing assignments are kept. Per object all rows are combined into one call.

    Behaviour per row:
      - Group already assigned the same way          -> Unchanged
      - App already assigned to the group with another intent, or include <-> exclude
        switched                                       -> the existing assignment is replaced (Changed)
      - Otherwise                                      -> Added

    Targets: group display name, group object ID, 'All Users' or 'All Devices'.
    Optional assignment filter (by name) with FilterMode include/exclude.

    CSV columns (delimiter -Delimiter, default ';'):
      Type;Name;Id;Group;Intent;Mode;Filter;FilterMode
        Name or Id        object display name or object ID
        Intent            apps only: required | available | uninstall (default required)
        Mode              include | exclude (default include; 'All Users/Devices' only include)
        Filter/FilterMode optional

.PARAMETER TenantId
    Tenant ID or primary domain.

.PARAMETER ClientId
    App registration for app-only auth (with -CertificateThumbprint). Omit for interactive login.

.PARAMETER CsvPath
    CSV with assignments. Alternative: -Type, -Name/-Id, -Group.

.PARAMETER Notifications
    For new Win32 app assignments: showAll | showReboot | hideAll (default showAll).

.NOTES
    Required Graph permissions (application or delegated):
      DeviceManagementApps.ReadWrite.All
      DeviceManagementConfiguration.ReadWrite.All
      DeviceManagementScripts.ReadWrite.All   (platform scripts)
      Group.Read.All
    Uses the beta endpoint for all Intune objects: settings catalog, admin templates,
    platform scripts and assignment filters are not (fully) available in v1.0.
    Remediations (deviceHealthScripts) are not supported: they need a run schedule per assignment.

    Exit codes: 0 = all rows OK, 1 = one or more rows failed/not found, 2 = fatal error.

.EXAMPLE
    .\Set-IntuneAssignments.ps1 -TenantId contoso.onmicrosoft.com -CsvPath .\assignments.csv -WhatIf

.EXAMPLE
    .\Set-IntuneAssignments.ps1 -TenantId contoso.onmicrosoft.com -Type App -Name '7-Zip','Google Chrome' `
        -Group 'SG-Intune-Devices-All' -Intent required -ClientId <appId> -CertificateThumbprint <thumb>
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Csv')]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$CsvPath,

    [Parameter(Mandatory, ParameterSetName = 'Params')]
    [ValidateSet('App', 'DeviceConfiguration', 'Compliance', 'SettingsCatalog', 'AdminTemplate', 'PlatformScript')]
    [string]$Type,

    [Parameter(ParameterSetName = 'Params')][string[]]$Name = @(),
    [Parameter(ParameterSetName = 'Params')][string[]]$Id = @(),
    [Parameter(Mandatory, ParameterSetName = 'Params')][string[]]$Group,
    [Parameter(ParameterSetName = 'Params')][ValidateSet('required', 'available', 'uninstall')][string]$Intent = 'required',
    [Parameter(ParameterSetName = 'Params')][ValidateSet('include', 'exclude')][string]$Mode = 'include',
    [Parameter(ParameterSetName = 'Params')][string]$Filter,
    [Parameter(ParameterSetName = 'Params')][ValidateSet('include', 'exclude')][string]$FilterMode = 'include',

    [ValidateSet('showAll', 'showReboot', 'hideAll')][string]$Notifications = 'showAll',

    [string]$OutputDirectory = (Get-Location).Path,
    [string]$Delimiter = ';'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Beta = 'https://graph.microsoft.com/beta'
$V1   = 'https://graph.microsoft.com/v1.0'

# Per type: collection URL, name property, body property of /assign
$TypeMap = @{
    App                 = @{ Path = 'deviceAppManagement/mobileApps';                 NameProp = 'displayName'; BodyProp = 'mobileAppAssignments' }
    DeviceConfiguration = @{ Path = 'deviceManagement/deviceConfigurations';          NameProp = 'displayName'; BodyProp = 'assignments' }
    Compliance          = @{ Path = 'deviceManagement/deviceCompliancePolicies';      NameProp = 'displayName'; BodyProp = 'assignments' }
    SettingsCatalog     = @{ Path = 'deviceManagement/configurationPolicies';         NameProp = 'name';        BodyProp = 'assignments' }
    AdminTemplate       = @{ Path = 'deviceManagement/groupPolicyConfigurations';     NameProp = 'displayName'; BodyProp = 'assignments' }
    PlatformScript      = @{ Path = 'deviceManagement/deviceManagementScripts';       NameProp = 'displayName'; BodyProp = 'deviceManagementScriptAssignments' }
}

#region Helpers ------------------------------------------------------------------------------

$script:LogFile = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false
    }
}

function Invoke-Graph {
    <# Invoke-MgGraphRequest with retry on 429/503/504 (honours Retry-After). #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        $Body
    )
    for ($attempt = 1; ; $attempt++) {
        try {
            $params = @{ Uri = $Uri; Method = $Method; OutputType = 'PSObject' }
            if ($null -ne $Body) {
                $params.Body        = ($Body | ConvertTo-Json -Depth 20)
                $params.ContentType = 'application/json'
            }
            return Invoke-MgGraphRequest @params
        }
        catch {
            $status = $null
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            if ($status -in 429, 503, 504 -and $attempt -le 5) {
                $wait = 10 * $attempt
                try {
                    $ra = $_.Exception.Response.Headers.RetryAfter.Delta
                    if ($ra) { $wait = [int]$ra.TotalSeconds }
                } catch { }
                Write-Log "  Graph $status, retry in $wait s ($attempt/5)" 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Get-GraphAll {
    <# GET with @odata.nextLink paging. #>
    param([Parameter(Mandatory)][string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = Invoke-Graph -Uri $next
        if ($resp.PSObject.Properties['value']) { foreach ($v in $resp.value) { $items.Add($v) } }
        $next = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
    }
    return ,$items
}

function Test-Guid([string]$Value) {
    $g = [guid]::Empty
    return [guid]::TryParse($Value, [ref]$g)
}

$script:GroupCache = @{}
function Resolve-Group([string]$Value) {
    if ($Value -match '^\s*All\s*Users\s*$')   { return @{ Kind = 'AllUsers';   Id = $null; Name = 'All Users' } }
    if ($Value -match '^\s*All\s*Devices\s*$') { return @{ Kind = 'AllDevices'; Id = $null; Name = 'All Devices' } }
    if ($script:GroupCache.ContainsKey($Value)) { return $script:GroupCache[$Value] }

    if (Test-Guid $Value) {
        $g = Invoke-Graph -Uri "$V1/groups/$($Value)?`$select=id,displayName"
    }
    else {
        $escaped = $Value.Replace("'", "''")
        $found = Get-GraphAll -Uri ("$V1/groups?`$filter=displayName eq '{0}'&`$select=id,displayName" -f [uri]::EscapeDataString($escaped))
        if ($found.Count -eq 0) { throw "Group '$Value' not found." }
        if ($found.Count -gt 1) { throw "Group name '$Value' is ambiguous ($($found.Count) groups); use the object ID." }
        $g = $found[0]
    }
    $r = @{ Kind = 'Group'; Id = $g.id; Name = $g.displayName }
    $script:GroupCache[$Value] = $r
    return $r
}

$script:FilterCache = $null
function Resolve-FilterId([string]$Value) {
    if (Test-Guid $Value) { return $Value }
    if (-not $script:FilterCache) {
        $script:FilterCache = Get-GraphAll -Uri "$Beta/deviceManagement/assignmentFilters?`$select=id,displayName"
    }
    $f = @($script:FilterCache | Where-Object { $_.displayName -eq $Value })
    if ($f.Count -ne 1) { throw "Assignment filter '$Value' not found or ambiguous." }
    return $f[0].id
}

$script:ObjectCache = @{}
function Resolve-IntuneObject([string]$ObjType, [string]$ObjName, [string]$ObjId) {
    $map = $TypeMap[$ObjType]
    if ($ObjId) {
        $o = Invoke-Graph -Uri "$Beta/$($map.Path)/$ObjId"
        $odataType = if ($o.PSObject.Properties['@odata.type']) { $o.'@odata.type' } else { $null }
        return [pscustomobject]@{ Id = $o.id; Name = $o.($map.NameProp); ODataType = $odataType }
    }
    if (-not $script:ObjectCache.ContainsKey($ObjType)) {
        $select = if ($ObjType -eq 'App') { '' } else { "?`$select=id,$($map.NameProp)" }
        $script:ObjectCache[$ObjType] = Get-GraphAll -Uri "$Beta/$($map.Path)$select"
    }
    $hits = @($script:ObjectCache[$ObjType] | Where-Object { $_.($map.NameProp) -eq $ObjName })
    if ($hits.Count -eq 0) { throw "$ObjType '$ObjName' not found." }
    if ($hits.Count -gt 1) { throw "$ObjType name '$ObjName' is ambiguous ($($hits.Count)); use the Id column." }
    $odataType = if ($hits[0].PSObject.Properties['@odata.type']) { $hits[0].'@odata.type' } else { $null }
    return [pscustomobject]@{ Id = $hits[0].id; Name = $hits[0].($map.NameProp); ODataType = $odataType }
}

function New-Target($GroupInfo, [string]$TargetMode, [string]$FilterId, [string]$FltMode) {
    $t = [ordered]@{}
    switch ($GroupInfo.Kind) {
        'AllUsers'   { $t['@odata.type'] = '#microsoft.graph.allLicensedUsersAssignmentTarget' }
        'AllDevices' { $t['@odata.type'] = '#microsoft.graph.allDevicesAssignmentTarget' }
        default {
            $t['@odata.type'] = if ($TargetMode -eq 'exclude') { '#microsoft.graph.exclusionGroupAssignmentTarget' }
                                else { '#microsoft.graph.groupAssignmentTarget' }
            $t['groupId'] = $GroupInfo.Id
        }
    }
    if ($FilterId -and $TargetMode -ne 'exclude') {
        $t['deviceAndAppManagementAssignmentFilterId']   = $FilterId
        $t['deviceAndAppManagementAssignmentFilterType'] = $FltMode
    }
    return $t
}

function Get-TargetKey($Target) {
    <# Identifies "who" an assignment targets, independent of include/exclude. #>
    $type = $Target.'@odata.type'
    if ($type -match 'allLicensedUsers') { return 'AllUsers' }
    if ($type -match 'allDevices')       { return 'AllDevices' }
    if ($Target.PSObject.Properties['groupId'] -or ($Target -is [System.Collections.IDictionary] -and $Target.Contains('groupId'))) {
        return "Group:$($Target.groupId)".ToLower()
    }
    return "Other:$type"
}

function ConvertTo-CleanTarget($Target) {
    <# Removes read-only/irrelevant properties from an existing target before re-posting. #>
    $t = [ordered]@{ '@odata.type' = $Target.'@odata.type' }
    foreach ($p in 'groupId', 'deviceAndAppManagementAssignmentFilterId', 'deviceAndAppManagementAssignmentFilterType') {
        if ($Target.PSObject.Properties[$p] -and $null -ne $Target.$p) { $t[$p] = $Target.$p }
    }
    if ($t.Contains('deviceAndAppManagementAssignmentFilterType') -and $t['deviceAndAppManagementAssignmentFilterType'] -eq 'none') {
        $t.Remove('deviceAndAppManagementAssignmentFilterId')
        $t.Remove('deviceAndAppManagementAssignmentFilterType')
    }
    return $t
}

function ConvertTo-CleanSettings($Settings) {
    if ($null -eq $Settings) { return $null }
    $s = [ordered]@{}
    foreach ($p in $Settings.PSObject.Properties) {
        if ($p.Name -ne 'id') { $s[$p.Name] = $p.Value }
    }
    return $s
}

#endregion

#region Main ---------------------------------------------------------------------------------

$results = New-Object System.Collections.Generic.List[object]
$exitCode = 0

try {
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile = Join-Path $OutputDirectory "IntuneAssignments_$($TenantId)_$stamp.log"

    # 1. Rows -----------------------------------------------------------------------------
    $rows = New-Object System.Collections.Generic.List[object]
    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV not found: $CsvPath" }
        foreach ($r in (Import-Csv -LiteralPath $CsvPath -Delimiter $Delimiter -Encoding UTF8)) {
            $get = { param($n, $d) if ($r.PSObject.Properties[$n] -and "$($r.$n)".Trim()) { "$($r.$n)".Trim() } else { $d } }
            $rows.Add([pscustomobject]@{
                Type       = & $get 'Type' ''
                Name       = & $get 'Name' ''
                Id         = & $get 'Id' ''
                Group      = & $get 'Group' ''
                Intent     = (& $get 'Intent' 'required').ToLower()
                Mode       = (& $get 'Mode' 'include').ToLower()
                Filter     = & $get 'Filter' ''
                FilterMode = (& $get 'FilterMode' 'include').ToLower()
            })
        }
    }
    else {
        if (-not $Name -and -not $Id) { throw 'Specify -Name or -Id.' }
        $objects = @($Name | ForEach-Object { @{ Name = $_; Id = '' } }) + @($Id | ForEach-Object { @{ Name = ''; Id = $_ } })
        foreach ($o in $objects) {
            foreach ($g in $Group) {
                $rows.Add([pscustomobject]@{
                    Type = $Type; Name = $o.Name; Id = $o.Id; Group = $g; Intent = $Intent
                    Mode = $Mode; Filter = "$Filter"; FilterMode = $FilterMode
                })
            }
        }
    }
    Write-Log "=== Intune bulk assignment for ${TenantId}: $($rows.Count) row(s)$(if ($WhatIfPreference) { ' (WhatIf)' }) ==="

    # 2. Connect --------------------------------------------------------------------------
    $scopes = 'DeviceManagementApps.ReadWrite.All', 'DeviceManagementConfiguration.ReadWrite.All',
              'DeviceManagementScripts.ReadWrite.All', 'Group.Read.All'
    if ($ClientId) {
        if (-not $CertificateThumbprint) { throw '-ClientId requires -CertificateThumbprint.' }
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    }
    else {
        Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome
    }

    # 3. Validate and resolve every row --------------------------------------------------
    $work = @{}   # key "Type|ObjectId" -> @{ Object; Type; Rows = list }
    foreach ($row in $rows) {
        $res = [ordered]@{
            Type = $row.Type; Object = if ($row.Name) { $row.Name } else { $row.Id }; Group = $row.Group
            Intent = $row.Intent; Mode = $row.Mode; Filter = $row.Filter; Status = ''; Detail = ''
        }
        try {
            if (-not $TypeMap.ContainsKey($row.Type)) { throw "Unknown Type '$($row.Type)'." }
            if (-not $row.Name -and -not $row.Id)     { throw 'Name or Id is required.' }
            if (-not $row.Group)                      { throw 'Group is required.' }
            if ($row.Mode -notin 'include', 'exclude') { throw "Invalid Mode '$($row.Mode)'." }
            if ($row.Type -eq 'App' -and $row.Intent -notin 'required', 'available', 'uninstall') {
                throw "Invalid Intent '$($row.Intent)'."
            }

            $obj = Resolve-IntuneObject $row.Type $row.Name $row.Id
            $grp = Resolve-Group $row.Group
            if ($grp.Kind -ne 'Group' -and $row.Mode -eq 'exclude') { throw "'$($grp.Name)' cannot be excluded." }
            $fid = if ($row.Filter) { Resolve-FilterId $row.Filter } else { $null }

            $res.Object = $obj.Name
            $res.Group  = $grp.Name
            $key = "$($row.Type)|$($obj.Id)"
            if (-not $work.ContainsKey($key)) {
                $work[$key] = @{ Object = $obj; Type = $row.Type; Items = New-Object System.Collections.Generic.List[object] }
            }
            $work[$key].Items.Add(@{ Row = $row; Group = $grp; FilterId = $fid; Result = $res })
        }
        catch {
            $res.Status = 'Failed'; $res.Detail = $_.Exception.Message
            Write-Log "Row $($res.Type) '$($res.Object)' -> '$($res.Group)': $($res.Detail)" 'ERROR'
            $results.Add([pscustomobject]$res)
            $exitCode = 1
        }
    }

    # 4. Merge and assign per object -----------------------------------------------------
    foreach ($entry in $work.Values) {
        $obj  = $entry.Object
        $map  = $TypeMap[$entry.Type]
        $base = "$Beta/$($map.Path)/$($obj.Id)"
        Write-Log "$($entry.Type) '$($obj.Name)':"

        try {
            $existing = Get-GraphAll -Uri "$base/assignments"

            # Existing assignments as an ordered list keyed by target
            $assignments = New-Object System.Collections.Specialized.OrderedDictionary
            foreach ($a in $existing) {
                $clean = [ordered]@{}
                if ($entry.Type -eq 'App') {
                    $clean['@odata.type'] = '#microsoft.graph.mobileAppAssignment'
                    $clean['intent']      = $a.intent
                    $settings = if ($a.PSObject.Properties['settings']) { ConvertTo-CleanSettings $a.settings } else { $null }
                    if ($null -ne $settings) { $clean['settings'] = $settings }
                }
                $clean['target'] = ConvertTo-CleanTarget $a.target
                $assignments[(Get-TargetKey $a.target)] = $clean
            }

            $changed = $false
            foreach ($item in $entry.Items) {
                $row = $item.Row; $res = $item.Result
                $target = New-Target $item.Group $row.Mode $item.FilterId $row.FilterMode
                $key    = Get-TargetKey ([pscustomobject]$target)

                $new = [ordered]@{}
                if ($entry.Type -eq 'App') {
                    $new['@odata.type'] = '#microsoft.graph.mobileAppAssignment'
                    $new['intent']      = $row.Intent
                    if ($obj.ODataType -eq '#microsoft.graph.win32LobApp') {
                        $new['settings'] = [ordered]@{
                            '@odata.type'                  = '#microsoft.graph.win32LobAppAssignmentSettings'
                            notifications                  = $Notifications
                            deliveryOptimizationPriority   = 'notConfigured'
                        }
                    }
                }
                $new['target'] = $target

                if ($assignments.Contains($key)) {
                    $old = $assignments[$key]
                    $same = ($old['target']['@odata.type'] -eq $target['@odata.type']) -and
                            ("$($old['target']['deviceAndAppManagementAssignmentFilterId'])" -eq "$($target['deviceAndAppManagementAssignmentFilterId'])") -and
                            ($entry.Type -ne 'App' -or $old['intent'] -eq $row.Intent)
                    if ($same) {
                        $res.Status = 'Unchanged'
                        Write-Log "  = $($item.Group.Name) already assigned"
                        continue
                    }
                    if ($entry.Type -eq 'App' -and $old.Contains('settings') -and $old['intent'] -eq $row.Intent) {
                        $new['settings'] = $old['settings']
                    }
                    $assignments[$key] = $new
                    $res.Status = 'Changed'
                    Write-Log "  ~ $($item.Group.Name): existing assignment replaced"
                }
                else {
                    $assignments[$key] = $new
                    $res.Status = 'Added'
                    Write-Log "  + $($item.Group.Name) ($($row.Mode)$(if ($entry.Type -eq 'App') { ", $($row.Intent)" }))"
                }
                $changed = $true
            }

            if ($changed) {
                $body = @{ $map.BodyProp = @($assignments.Values) }
                if ($PSCmdlet.ShouldProcess("$($entry.Type) '$($obj.Name)'", "Set $($assignments.Count) assignment(s)")) {
                    $null = Invoke-Graph -Uri "$base/assign" -Method POST -Body $body
                    Write-Log "  Assigned ($($assignments.Count) total)."
                }
                else {
                    foreach ($item in $entry.Items) {
                        if ($item.Result.Status -in 'Added', 'Changed') { $item.Result.Status = "WhatIf:$($item.Result.Status)" }
                    }
                }
            }
        }
        catch {
            $msg = $_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $msg = $_.ErrorDetails.Message }
            Write-Log "  Failed: $msg" 'ERROR'
            foreach ($item in $entry.Items) {
                if ($item.Result.Status -ne 'Unchanged') { $item.Result.Status = 'Failed'; $item.Result.Detail = $msg }
            }
            $exitCode = 1
        }

        foreach ($item in $entry.Items) { $results.Add([pscustomobject]$item.Result) }
    }

    # 5. Report ---------------------------------------------------------------------------
    $csvOut = Join-Path $OutputDirectory "IntuneAssignments_$($TenantId)_$stamp.csv"
    $results | Export-Csv -LiteralPath $csvOut -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter -WhatIf:$false

    Write-Host ''
    $results | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Log ("{0,-18} {1}" -f $_.Name, $_.Count) }
    Write-Log "Result: $csvOut"
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}

exit $exitCode

#endregion
