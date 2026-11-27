#Requires -Version 5.1

<#
.SYNOPSIS
    Licence overview for one tenant: subscriptions, assignments per user and users without a licence.

.DESCRIPTION
    Produces three CSV files in -OutputDirectory:

      Licenses_Skus_<domain>_<date>.csv         per SKU: enabled, consumed, unused, suspended,
                                                warning and locked-out units, with friendly names
      Licenses_Assignments_<domain>_<date>.csv  per user per SKU: direct and/or group-based
                                                (group names), assignment state and errors
      Licenses_Unlicensed_<domain>_<date>.csv   users without any licence (default: enabled
                                                member accounts only)

    Friendly product names come from Microsoft's "Product names and service plan identifiers
    for licensing" CSV (cached in %TEMP% for 7 days). If the download fails, the
    skuPartNumber is used.

    SKUs with 10,000 or more enabled units are free/viral offers (e.g. Power Automate Free,
    Teams Exploratory). They are marked IsFreeOrViral and left out of the "unused" findings.

.PARAMETER TenantId
    Tenant ID or initial domain (contoso.onmicrosoft.com).

.PARAMETER ClientId
    App registration (client) ID for app-only authentication. Without it the script signs in
    interactively (delegated).

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate (CurrentUser\My or LocalMachine\My) for app-only authentication.

.PARAMETER IncludeDisabled
    Also list disabled accounts in the unlicensed report.

.PARAMETER IncludeGuests
    Also list guest accounts in the unlicensed report.

.PARAMETER OutputDirectory
    Folder for CSVs and log. Default: current directory.

.PARAMETER Delimiter
    CSV delimiter. Default ';' (Dutch Excel).

.EXAMPLE
    .\Get-LicenseOverview.ps1 -TenantId contoso.onmicrosoft.com

.EXAMPLE
    .\Get-LicenseOverview.ps1 -TenantId $tid -ClientId $appId -CertificateThumbprint $thumb -OutputDirectory C:\Reports

.NOTES
    Required Graph permissions (application or delegated):
      Organization.Read.All   subscribedSkus and tenant name
      User.Read.All           users, assignedLicenses, licenseAssignmentStates
      Group.Read.All          names of licensing groups
    (Directory.Read.All covers all three.)
    Delegated: the signed-in user also needs a role such as Global Reader or License Administrator.

    Dependency: Microsoft.Graph.Authentication.
    Exit codes: 0 = no findings, 1 = findings (unused paid licences, unlicensed users,
    licensed disabled accounts or assignment errors), 2 = fatal error.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,

    [switch]$IncludeDisabled,
    [switch]$IncludeGuests,

    [string]$OutputDirectory = (Get-Location).Path,
    [string]$Delimiter = ';'
)

# StrictMode 1 only: Graph objects may omit properties, and level 2+ throws on that.
Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

$ProductNamesUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
$Scopes = @('Organization.Read.All', 'User.Read.All', 'Group.Read.All')

#region Shared Graph helpers (identical in all tenant scripts) --------------------------------

$script:LogFile = $null
$script:Tenant  = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
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

function Initialize-Output([string]$Prefix) {
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $OutputDirectory ('{0}_{1:yyyyMMdd_HHmmss}.log' -f $Prefix, (Get-Date))
}

function Connect-Tenant([string[]]$Scopes) {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Module Microsoft.Graph.Authentication is missing: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if ($ClientId) {
        if (-not $CertificateThumbprint) { throw '-CertificateThumbprint is required together with -ClientId.' }
        Write-Log "Connecting to $TenantId (app-only, client $ClientId)"
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
    }
    else {
        Write-Log "Connecting to $TenantId (delegated: $($Scopes -join ', '))"
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome -ErrorAction Stop
    }
}

function Get-HttpStatus($ErrorRecord) {
    try { $code = [int]$ErrorRecord.Exception.Response.StatusCode; if ($code) { return $code } } catch { }
    if ("$($ErrorRecord.Exception.Message)" -match '\b(429|503|504)\b') { return [int]$Matches[1] }
    return 0
}

function Get-RetryDelay($ErrorRecord, [int]$Attempt) {
    try {
        $ra = $ErrorRecord.Exception.Response.Headers.RetryAfter
        if ($ra -and $ra.Delta) { return [int][math]::Ceiling($ra.Delta.TotalSeconds) }
    } catch { }
    return [int][math]::Min(60, [math]::Pow(2, $Attempt))
}

function Invoke-Graph {
    <#
        Calls Graph. For GET collections it follows @odata.nextLink and emits the items.
        Retries on 429/503/504 and honours Retry-After.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string]$Method = 'GET',
        $Body,
        [int]$MaxRetries = 6
    )
    $next = $Uri
    while ($next) {
        $attempt = 0
        while ($true) {
            try {
                $req = @{ Method = $Method; Uri = $next; OutputType = 'PSObject'; ErrorAction = 'Stop' }
                if ($null -ne $Body) {
                    $req.Body = ($Body | ConvertTo-Json -Depth 20)
                    $req.ContentType = 'application/json'
                }
                $resp = Invoke-MgGraphRequest @req
                break
            }
            catch {
                $status = Get-HttpStatus $_
                if ($status -in 429, 503, 504 -and $attempt -lt $MaxRetries) {
                    $attempt++
                    $delay = Get-RetryDelay $_ $attempt
                    Write-Log "  Graph returned $status; retry $attempt in $delay s" 'WARN'
                    Start-Sleep -Seconds $delay
                    continue
                }
                throw
            }
        }
        if ($Method -ne 'GET' -or $null -eq $resp -or -not $resp.PSObject.Properties['value']) { return $resp }
        foreach ($item in $resp.value) { $item }
        $nl = $resp.PSObject.Properties['@odata.nextLink']
        $next = if ($nl) { $nl.Value } else { $null }
    }
}

function Get-TenantLabel {
    try {
        $org = @(Invoke-Graph -Uri 'v1.0/organization?$select=displayName,verifiedDomains')[0]
        $dom = @($org.verifiedDomains | Where-Object { $_.isDefault }) + @($org.verifiedDomains | Where-Object { $_.isInitial })
        return [pscustomobject]@{ Name = $org.displayName; Domain = $dom[0].name }
    }
    catch {
        Write-Log 'Could not read organization (Organization.Read.All); using the tenant ID as label.' 'WARN'
        return [pscustomobject]@{ Name = $TenantId; Domain = $TenantId }
    }
}

function ConvertTo-DateTimeUtc($Value) {
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    return [datetime]::Parse("$Value", [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal)
}

function Format-LocalDate($Value) {
    $d = ConvertTo-DateTimeUtc $Value
    if ($d) { $d.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
}

function Export-Report($Rows, [string]$Name) {
    $path = Join-Path $OutputDirectory ('{0}_{1}_{2:yyyyMMdd}.csv' -f $Name, $script:Tenant.Domain, (Get-Date))
    $list = @($Rows | Where-Object { $null -ne $_ })
    $list | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter -WhatIf:$false
    Write-Log "Exported $($list.Count) row(s) to $path"
    return $path
}

#endregion

#region Script-specific helpers -----------------------------------------------------------------

function Get-ProductNameMap {
    $map = @{}
    $cache = Join-Path $env:TEMP 'MsLicensingProductNames.csv'
    try {
        if (-not (Test-Path -LiteralPath $cache) -or (Get-Item -LiteralPath $cache).LastWriteTime -lt (Get-Date).AddDays(-7)) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $ProductNamesUrl -OutFile $cache -UseBasicParsing -ErrorAction Stop
        }
        foreach ($r in (Import-Csv -LiteralPath $cache -Encoding UTF8)) {
            if ($r.GUID -and -not $map.ContainsKey($r.GUID)) { $map[$r.GUID] = $r.Product_Display_Name }
        }
        Write-Log "Friendly product names loaded: $($map.Count) SKUs"
    }
    catch {
        Write-Log "Could not load product names ($($_.Exception.Message)); using skuPartNumber." 'WARN'
    }
    return $map
}

$script:GroupNames = @{}
function Get-GroupName([string]$Id) {
    if (-not $script:GroupNames.ContainsKey($Id)) {
        try   { $script:GroupNames[$Id] = (Invoke-Graph -Uri "v1.0/groups/${Id}?`$select=displayName").displayName }
        catch { $script:GroupNames[$Id] = $Id }
    }
    return $script:GroupNames[$Id]
}

#endregion

#region Main ---------------------------------------------------------------------------------

$exitCode = 2
try {
    Initialize-Output 'Licenses'
    Connect-Tenant -Scopes $Scopes
    $script:Tenant = Get-TenantLabel
    Write-Log "=== Licence overview for $($script:Tenant.Name) ($($script:Tenant.Domain)) ==="

    $names = Get-ProductNameMap

    # 1. Subscriptions ------------------------------------------------------------------------
    $skuName = @{}
    $skuPart = @{}
    $skuRows = foreach ($s in @(Invoke-Graph -Uri 'v1.0/subscribedSkus')) {
        $id       = "$($s.skuId)"
        $friendly = if ($names.ContainsKey($id) -and $names[$id]) { $names[$id] } else { $s.skuPartNumber }
        $skuName[$id] = $friendly
        $skuPart[$id] = $s.skuPartNumber

        $enabled  = [int]$s.prepaidUnits.enabled
        $consumed = [int]$s.consumedUnits
        [pscustomobject][ordered]@{
            Product          = $friendly
            SkuPartNumber    = $s.skuPartNumber
            SkuId            = $id
            AppliesTo        = $s.appliesTo
            CapabilityStatus = $s.capabilityStatus
            Enabled          = $enabled
            Consumed         = $consumed
            Unused           = [math]::Max(0, $enabled - $consumed)
            Suspended        = [int]$s.prepaidUnits.suspended
            Warning          = [int]$s.prepaidUnits.warning
            LockedOut        = [int]$s.prepaidUnits.lockedOut
            OverAssigned     = ($consumed -gt $enabled)
            IsFreeOrViral    = ($enabled -ge 10000)
        }
    }
    $skuRows = @($skuRows)
    Write-Log "Subscriptions: $($skuRows.Count)"

    # 2. Users and assignments ----------------------------------------------------------------
    $users = @(Invoke-Graph -Uri 'v1.0/users?$select=id,userPrincipalName,displayName,accountEnabled,userType,assignedLicenses,licenseAssignmentStates,department,usageLocation&$top=999')
    Write-Log "Users: $($users.Count)"

    $assignRows = New-Object System.Collections.Generic.List[object]
    foreach ($u in $users) {
        # Union of assigned SKUs and SKUs in assignment states (an errored group assignment
        # appears only in licenseAssignmentStates)
        $skuIds = New-Object System.Collections.Generic.List[string]
        foreach ($l in $u.assignedLicenses)        { if (-not $skuIds.Contains("$($l.skuId)")) { $skuIds.Add("$($l.skuId)") } }
        foreach ($st in $u.licenseAssignmentStates) { if (-not $skuIds.Contains("$($st.skuId)")) { $skuIds.Add("$($st.skuId)") } }

        foreach ($sid in $skuIds) {
            $direct = $false
            $groups = New-Object System.Collections.Generic.List[string]
            $states = New-Object System.Collections.Generic.List[string]
            $errors = New-Object System.Collections.Generic.List[string]
            foreach ($st in $u.licenseAssignmentStates) {
                if ("$($st.skuId)" -ne $sid) { continue }
                if ($st.assignedByGroup) { $groups.Add((Get-GroupName $st.assignedByGroup)) } else { $direct = $true }
                if ($st.state -and -not $states.Contains($st.state)) { $states.Add($st.state) }
                if ($st.error -and $st.error -ne 'None') { $errors.Add($st.error) }
            }
            $disabledPlans = 0
            foreach ($l in $u.assignedLicenses) { if ("$($l.skuId)" -eq $sid) { $disabledPlans = @($l.disabledPlans | Where-Object { $_ }).Count } }

            $type = @()
            if ($direct)           { $type += 'Direct' }
            if ($groups.Count -gt 0) { $type += 'Group' }

            $assignRows.Add([pscustomobject][ordered]@{
                UserPrincipalName = $u.userPrincipalName
                DisplayName       = $u.displayName
                AccountEnabled    = $u.accountEnabled
                UserType          = $u.userType
                Department        = $u.department
                UsageLocation     = $u.usageLocation
                Product           = $(if ($skuName.ContainsKey($sid)) { $skuName[$sid] } else { $sid })
                SkuPartNumber     = $skuPart[$sid]
                AssignmentType    = ($type -join '+')
                AssignedByGroups  = ($groups -join '; ')
                DisabledPlans     = $disabledPlans
                State             = ($states -join ',')
                Error             = ($errors -join ',')
            })
        }
    }

    # 3. Unlicensed users -----------------------------------------------------------------------
    $unlicensed = @($users | Where-Object {
            (-not $_.assignedLicenses) -and
            ($IncludeGuests -or $_.userType -ne 'Guest') -and
            ($IncludeDisabled -or $_.accountEnabled)
        } | ForEach-Object {
            [pscustomobject][ordered]@{
                UserPrincipalName = $_.userPrincipalName
                DisplayName       = $_.displayName
                AccountEnabled    = $_.accountEnabled
                UserType          = $_.userType
                Department        = $_.department
                UsageLocation     = $_.usageLocation
            }
        })

    # 4. Export -----------------------------------------------------------------------------------
    $null = Export-Report ($skuRows | Sort-Object Product) 'Licenses_Skus'
    $null = Export-Report ($assignRows | Sort-Object UserPrincipalName, Product) 'Licenses_Assignments'
    $null = Export-Report ($unlicensed | Sort-Object UserPrincipalName) 'Licenses_Unlicensed'

    # 5. Summary ----------------------------------------------------------------------------------
    $paidUnused     = @($skuRows | Where-Object { -not $_.IsFreeOrViral -and $_.Unused -gt 0 -and $_.CapabilityStatus -eq 'Enabled' })
    $overAssigned   = @($skuRows | Where-Object { $_.OverAssigned })
    $errorRows      = @($assignRows | Where-Object { $_.Error })
    $disabledLic    = @($assignRows | Where-Object { $_.AccountEnabled -eq $false } | Select-Object -ExpandProperty UserPrincipalName -Unique)

    Write-Host ''
    Write-Host "Subscriptions ($($script:Tenant.Domain)):" -ForegroundColor Cyan
    $skuRows | Where-Object { -not $_.IsFreeOrViral } | Sort-Object Product |
        Format-Table Product, Enabled, Consumed, Unused, Suspended, Warning, CapabilityStatus -AutoSize | Out-String | Write-Host

    $summary = ("Summary: {0} paid SKU(s) with unused units ({1} units in total), {2} over-assigned, " +
        "{3} unlicensed user(s), {4} disabled account(s) still licensed, {5} assignment error(s).") -f `
        $paidUnused.Count, (($paidUnused | Measure-Object Unused -Sum).Sum + 0), $overAssigned.Count,
        $unlicensed.Count, $disabledLic.Count, $errorRows.Count
    Write-Log $summary

    $findings = $paidUnused.Count + $overAssigned.Count + $unlicensed.Count + $disabledLic.Count + $errorRows.Count
    $exitCode = if ($findings -gt 0) { 1 } else { 0 }
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
