#Requires -Version 5.1

<#
.SYNOPSIS
    Reports Autopilot devices that have no (successfully) assigned deployment profile.

.DESCRIPTION
    Fetches all windowsAutopilotDeviceIdentities and reports every device whose
    deploymentProfileAssignmentStatus is not "assigned...": notAssigned, pending, failed,
    unknown. With -IncludeOutOfSync, assignedOutOfSync is reported as well.

    For each reported device the script also reads the assigned and the intended profile
    ($expand=deploymentProfile,intendedDeploymentProfile), so you can see which profile is
    on its way for a pending device. Skip this with -SkipProfileLookup on large sets.

    Uses the BETA endpoint: deploymentProfileAssignmentStatus, the detailed status,
    deploymentProfile and intendedDeploymentProfile do not exist in v1.0.

.PARAMETER TenantId
    Tenant ID or tenant domain.

.PARAMETER ClientId
    App registration (client) ID for app-only sign-in with a certificate. Without
    -ClientId the script signs in interactively (delegated).

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate (CurrentUser\My or LocalMachine\My) for -ClientId.

.PARAMETER IncludeOutOfSync
    Also report devices with status assignedOutOfSync.

.PARAMETER SkipProfileLookup
    Skips the per-device lookup of the assigned/intended profile name.

.PARAMETER OutputDirectory
    Folder for the CSV and the log. Default: current folder.

.PARAMETER Delimiter
    CSV delimiter. Default ';' (Dutch Excel).

.EXAMPLE
    .\Get-AutopilotDevicesWithoutProfile.ps1 -TenantId contoso.onmicrosoft.com

.EXAMPLE
    .\Get-AutopilotDevicesWithoutProfile.ps1 -TenantId $tid -ClientId $app -CertificateThumbprint $thumb -IncludeOutOfSync

.NOTES
    Graph permissions
      Application : DeviceManagementServiceConfig.Read.All
                    Organization.Read.All (optional, for the tenant domain in file names)
      Delegated   : DeviceManagementServiceConfig.Read.All, User.Read

    Dependency: Microsoft.Graph.Authentication.
    A freshly imported device usually stays 'pending' or 'notAssigned' for up to
    15-30 minutes after its profile group membership is evaluated.

    Exit codes: 0 = all devices have a profile, 1 = devices without a profile found,
    2 = fatal error.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$IncludeOutOfSync,
    [switch]$SkipProfileLookup,
    [string]$OutputDirectory = (Get-Location).Path,
    [string]$Delimiter = ';'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Common helpers (identical in all tenant scripts) --------------------------------------

$script:LogFile     = $null
$script:LogBuffer   = New-Object System.Collections.Generic.List[string]
$script:TenantLabel = $TenantId

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
    if ($script:LogFile) { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false }
    else { $script:LogBuffer.Add($line) }
}

function Initialize-Output {
    param([Parameter(Mandatory)][string]$ScriptName)
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $OutputDirectory ('{0}_{1}_{2:yyyyMMdd_HHmmss}.log' -f $ScriptName, $script:TenantLabel, (Get-Date))
    if ($script:LogBuffer.Count -gt 0) {
        Add-Content -LiteralPath $script:LogFile -Value $script:LogBuffer -Encoding UTF8 -WhatIf:$false
        $script:LogBuffer.Clear()
    }
}

function Get-Value {
    <# Safe property access under StrictMode: $null when the property does not exist. #>
    param($Object, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function ConvertTo-LocalDate {
    <# Graph dates arrive as string or DateTime; '0001-01-01' means "never". #>
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    $d = if ($Value -is [datetime]) { $Value }
         else { [datetime]::Parse("$Value", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind) }
    if ($d.Year -le 1601) { return $null }
    return $d.ToLocalTime()
}

function Connect-Tenant {
    param([Parameter(Mandatory)][string[]]$Scopes)
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Module Microsoft.Graph.Authentication is missing: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    # ContextScope Process: no token reuse between runs against different tenants
    if ($ClientId) {
        if (-not $CertificateThumbprint) { throw '-CertificateThumbprint is required together with -ClientId.' }
        Write-Log "Connecting to $TenantId (app-only, client $ClientId)"
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint `
            -ContextScope Process -NoWelcome
    }
    else {
        Write-Log "Connecting to $TenantId (delegated, scopes: $($Scopes -join ', '))"
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -ContextScope Process -NoWelcome
    }
}

function Invoke-GraphApi {
    <# Invoke-MgGraphRequest with retry on 429/503/504, honoring Retry-After. #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        $Body,
        [int]$MaxRetries = 6
    )
    for ($attempt = 1; ; $attempt++) {
        try {
            $req = @{ Uri = $Uri; Method = $Method; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($null -ne $Body) {
                $req.Body        = ($Body | ConvertTo-Json -Depth 10 -Compress)
                $req.ContentType = 'application/json'
            }
            return Invoke-MgGraphRequest @req
        }
        catch {
            $status = $null
            $wait   = $null
            $respProp = $_.Exception.PSObject.Properties['Response']
            if ($respProp -and $respProp.Value) {
                try { $status = [int]$respProp.Value.StatusCode } catch { }
                try {
                    $ra = $respProp.Value.Headers.RetryAfter
                    if ($ra -and $ra.Delta) { $wait = [int][math]::Ceiling($ra.Delta.TotalSeconds) }
                } catch { }
            }
            if (-not $status -and $_.Exception.Message -match '\b(429|503|504)\b') { $status = [int]$Matches[1] }

            if ($status -in 429, 503, 504 -and $attempt -le $MaxRetries) {
                if (-not $wait) { $wait = [int][math]::Min(60, [math]::Pow(2, $attempt)) }
                Write-Log "  Graph returned $status; retry $attempt/$MaxRetries in $wait s" 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Get-GraphAll {
    <# GET with @odata.nextLink paging; returns all items of 'value'. #>
    param([Parameter(Mandatory)][string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = Invoke-GraphApi -Uri $next
        $val = Get-Value $resp 'value'
        if ($null -ne $val) { foreach ($v in @($val)) { $items.Add($v) } }
        $next = Get-Value $resp '@odata.nextLink'
    }
    return , $items.ToArray()
}

function Get-TenantLabel {
    try {
        $org = Get-GraphAll -Uri 'v1.0/organization?$select=verifiedDomains'
        $domain = @(Get-Value @($org)[0] 'verifiedDomains') |
            Where-Object { Get-Value $_ 'isDefault' } | Select-Object -First 1
        if ($domain) { return ((Get-Value $domain 'name') -replace '[\\/:*?"<>|]', '_') }
    }
    catch { Write-Log 'Could not read the tenant domain (Organization.Read.All missing?); using the tenant ID in file names.' 'WARN' }
    return ($TenantId -replace '[\\/:*?"<>|]', '_')
}

function Export-Report {
    param([AllowNull()][object[]]$Data, [Parameter(Mandatory)][string]$Name)
    $path = Join-Path $OutputDirectory ('{0}_{1}_{2:yyyyMMdd}.csv' -f $Name, $script:TenantLabel, (Get-Date))
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -gt 0) {
        $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter -WhatIf:$false
    }
    else {
        Set-Content -LiteralPath $path -Value '' -Encoding UTF8 -WhatIf:$false
    }
    Write-Log "Exported $($rows.Count) row(s) to $path"
    return $path
}

#endregion

#region Script ------------------------------------------------------------------------------

$exitCode = 0
try {
    Connect-Tenant -Scopes 'DeviceManagementServiceConfig.Read.All', 'User.Read'
    $script:TenantLabel = Get-TenantLabel
    Initialize-Output -ScriptName 'AutopilotNoProfile'
    Write-Log "=== Autopilot devices without profile for $script:TenantLabel ==="

    # beta: assignment status fields and profile navigation properties are not in v1.0
    $base = 'beta/deviceManagement/windowsAutopilotDeviceIdentities'
    $all  = Get-GraphAll -Uri $base
    Write-Log "$($all.Count) Autopilot device(s) in the tenant"

    $okStatuses = @('assignedInSync', 'assignedUnkownSyncState', 'assignedUnknownSyncState')
    if (-not $IncludeOutOfSync) { $okStatuses += 'assignedOutOfSync' }

    $flagged = @($all | Where-Object { "$(Get-Value $_ 'deploymentProfileAssignmentStatus')" -notin $okStatuses })
    Write-Log "$($flagged.Count) device(s) without a (synced) profile"

    $rows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($d in $flagged) {
        $i++
        $id = Get-Value $d 'id'
        $assignedProfile = $null
        $intendedProfile = $null

        if (-not $SkipProfileLookup) {
            if ($i % 25 -eq 0) { Write-Log "  Profile lookup $i/$($flagged.Count)" }
            try {
                $detail = Invoke-GraphApi -Uri "$base/$id`?`$expand=deploymentProfile,intendedDeploymentProfile"
                $assignedProfile = Get-Value (Get-Value $detail 'deploymentProfile') 'displayName'
                $intendedProfile = Get-Value (Get-Value $detail 'intendedDeploymentProfile') 'displayName'
            }
            catch { Write-Log "  Profile lookup failed for $(Get-Value $d 'serialNumber'): $($_.Exception.Message)" 'WARN' }
        }

        $lastContact = ConvertTo-LocalDate (Get-Value $d 'lastContactedDateTime')
        $rows.Add([pscustomobject]@{
            SerialNumber          = Get-Value $d 'serialNumber'
            Manufacturer          = Get-Value $d 'manufacturer'
            Model                 = Get-Value $d 'model'
            GroupTag              = Get-Value $d 'groupTag'
            PurchaseOrder         = Get-Value $d 'purchaseOrderIdentifier'
            ProfileStatus         = Get-Value $d 'deploymentProfileAssignmentStatus'
            ProfileDetailedStatus = Get-Value $d 'deploymentProfileAssignmentDetailedStatus'
            ProfileAssignedOn     = ConvertTo-LocalDate (Get-Value $d 'deploymentProfileAssignedDateTime')
            AssignedProfile       = $assignedProfile
            IntendedProfile       = $intendedProfile
            EnrollmentState       = Get-Value $d 'enrollmentState'
            LastContacted         = $lastContact
            AssignedUser          = Get-Value $d 'userPrincipalName'
            EntraDeviceId         = Get-Value $d 'azureActiveDirectoryDeviceId'
            IntuneDeviceId        = Get-Value $d 'managedDeviceId'
            AutopilotId           = $id
        })
    }

    $null = Export-Report -Data ($rows | Sort-Object ProfileStatus, GroupTag, SerialNumber) -Name 'AutopilotNoProfile'

    Write-Host ''
    Write-Log "Summary: $($all.Count) Autopilot device(s), $($rows.Count) without a (synced) profile"
    $rows | Group-Object ProfileStatus | Sort-Object Count -Descending | ForEach-Object {
        Write-Log ('  {0,-26} {1}' -f $_.Name, $_.Count)
    }
    $rows | Group-Object GroupTag | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object {
        Write-Log ('  group tag {0,-16} {1}' -f $(if ($_.Name) { "'$($_.Name)'" } else { '(none)' }), $_.Count)
    }
    if ($rows.Count -gt 0) { $exitCode = 1 }
}
catch {
    if (-not $script:LogFile) { try { Initialize-Output -ScriptName 'AutopilotNoProfile' } catch { } }
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
exit $exitCode

#endregion
