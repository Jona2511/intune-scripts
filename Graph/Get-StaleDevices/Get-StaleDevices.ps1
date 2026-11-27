#Requires -Version 5.1

<#
.SYNOPSIS
    Reports Intune and Entra ID devices that have not checked in for more than X days.

.DESCRIPTION
    Report only: nothing is disabled or deleted.

      - Intune  : managedDevices with lastSyncDateTime older than -Days.
      - Entra ID: devices with approximateLastSignInDateTime older than -Days. A device
                  that never signed in counts as stale when it was registered before
                  the cutoff.

    Both sources are matched on Intune azureADDeviceId = Entra deviceId. The combined
    report shows one row per device with StaleIn = Intune, Entra or Both, and
    Sources = Intune+Entra, IntuneOnly or EntraOnly.

    Entra objects created by an Autopilot import (physicalIds contains [ZTDID]) are
    marked IsAutopilotObject = True. Do not delete those as stale: they are needed for
    future Autopilot deployments.

    Output: three CSVs (Intune, Entra, combined) plus a log.

.PARAMETER TenantId
    Tenant ID or tenant domain.

.PARAMETER Days
    Number of days without check-in before a device is stale (default 90).

.PARAMETER ClientId
    App registration (client) ID for app-only sign-in with a certificate. Without
    -ClientId the script signs in interactively (delegated).

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate (CurrentUser\My or LocalMachine\My) for -ClientId.

.PARAMETER OutputDirectory
    Folder for the CSVs and the log. Default: current folder.

.PARAMETER Delimiter
    CSV delimiter. Default ';' (Dutch Excel).

.EXAMPLE
    .\Get-StaleDevices.ps1 -TenantId contoso.onmicrosoft.com

.EXAMPLE
    .\Get-StaleDevices.ps1 -TenantId $tid -ClientId $app -CertificateThumbprint $thumb -Days 120 -OutputDirectory D:\Reports

.NOTES
    Graph permissions
      Application : DeviceManagementManagedDevices.Read.All, Device.Read.All
                    Organization.Read.All (optional, for the tenant domain in file names)
      Delegated   : DeviceManagementManagedDevices.Read.All, Device.Read.All, User.Read

    Dependency: Microsoft.Graph.Authentication.
    approximateLastSignInDateTime is only updated about once every 14 days, so Entra
    dates can lag by up to two weeks.
    PrimaryUserUPN comes from managedDevice.userPrincipalName (the Intune primary user).

    Exit codes: 0 = no stale devices, 1 = stale devices found, 2 = fatal error.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [ValidateRange(1, 3650)][int]$Days = 90,
    [string]$ClientId,
    [string]$CertificateThumbprint,
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

function Get-DaysSince($Date) {
    if ($null -eq $Date) { return $null }
    return [int][math]::Floor(((Get-Date) - $Date).TotalDays)
}

$exitCode = 0
try {
    Connect-Tenant -Scopes 'DeviceManagementManagedDevices.Read.All', 'Device.Read.All', 'User.Read'
    $script:TenantLabel = Get-TenantLabel
    Initialize-Output -ScriptName 'StaleDevices'
    $cutoff = (Get-Date).AddDays(-$Days)
    Write-Log ("=== Stale devices for {0}: no check-in since {1:yyyy-MM-dd} ({2} days) ===" -f $script:TenantLabel, $cutoff, $Days)

    # 1. Intune --------------------------------------------------------------------------
    $intuneSelect = 'id,deviceName,operatingSystem,osVersion,managedDeviceOwnerType,lastSyncDateTime,' +
                    'enrolledDateTime,azureADDeviceId,serialNumber,userPrincipalName,complianceState,model,manufacturer'
    $intune = Get-GraphAll -Uri "v1.0/deviceManagement/managedDevices?`$select=$intuneSelect"
    Write-Log "Intune: $($intune.Count) managed device(s)"

    # 2. Entra ID -----------------------------------------------------------------------
    $entraSelect = 'id,deviceId,displayName,operatingSystem,operatingSystemVersion,approximateLastSignInDateTime,' +
                   'accountEnabled,trustType,profileType,registrationDateTime,isManaged,isCompliant,physicalIds'
    $entra = Get-GraphAll -Uri "v1.0/devices?`$select=$entraSelect&`$top=999"
    Write-Log "Entra ID: $($entra.Count) device(s)"

    # 3. Normalize ------------------------------------------------------------------------
    $intuneRows = foreach ($m in $intune) {
        $last = ConvertTo-LocalDate (Get-Value $m 'lastSyncDateTime')
        $enrolled = ConvertTo-LocalDate (Get-Value $m 'enrolledDateTime')
        $stale = if ($last) { $last -lt $cutoff } elseif ($enrolled) { $enrolled -lt $cutoff } else { $true }
        [pscustomobject]@{
            DeviceName      = Get-Value $m 'deviceName'
            OperatingSystem = Get-Value $m 'operatingSystem'
            OSVersion       = Get-Value $m 'osVersion'
            Ownership       = Get-Value $m 'managedDeviceOwnerType'
            IntuneLastSync  = $last
            DaysSinceSync   = Get-DaysSince $last
            EnrolledOn      = $enrolled
            SerialNumber    = Get-Value $m 'serialNumber'
            Manufacturer    = Get-Value $m 'manufacturer'
            Model           = Get-Value $m 'model'
            PrimaryUserUPN  = Get-Value $m 'userPrincipalName'
            ComplianceState = Get-Value $m 'complianceState'
            AzureADDeviceId = Get-Value $m 'azureADDeviceId'
            IntuneDeviceId  = Get-Value $m 'id'
            IsStale         = $stale
        }
    }
    $intuneRows = @($intuneRows)

    $entraRows = foreach ($e in $entra) {
        $last = ConvertTo-LocalDate (Get-Value $e 'approximateLastSignInDateTime')
        $reg  = ConvertTo-LocalDate (Get-Value $e 'registrationDateTime')
        $stale = if ($last) { $last -lt $cutoff } elseif ($reg) { $reg -lt $cutoff } else { $true }
        $isAutopilot = [bool](@(Get-Value $e 'physicalIds') | Where-Object { "$_" -like '`[ZTDID`]*' })
        [pscustomobject]@{
            DisplayName       = Get-Value $e 'displayName'
            OperatingSystem   = Get-Value $e 'operatingSystem'
            OSVersion         = Get-Value $e 'operatingSystemVersion'
            EntraLastSignIn   = $last
            DaysSinceSignIn   = Get-DaysSince $last
            RegisteredOn      = $reg
            AccountEnabled    = Get-Value $e 'accountEnabled'
            TrustType         = Get-Value $e 'trustType'
            ProfileType       = Get-Value $e 'profileType'
            IsManaged         = Get-Value $e 'isManaged'
            IsCompliant       = Get-Value $e 'isCompliant'
            IsAutopilotObject = $isAutopilot
            DeviceId          = Get-Value $e 'deviceId'
            ObjectId          = Get-Value $e 'id'
            IsStale           = $stale
        }
    }
    $entraRows = @($entraRows)

    # 4. Combine -------------------------------------------------------------------------
    $entraById = @{}
    foreach ($e in $entraRows) { if ($e.DeviceId) { $entraById[$e.DeviceId] = $e } }
    $matchedEntra = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $emptyGuid = '00000000-0000-0000-0000-000000000000'

    $combined = New-Object System.Collections.Generic.List[object]
    foreach ($i in $intuneRows) {
        $e = $null
        if ($i.AzureADDeviceId -and $i.AzureADDeviceId -ne $emptyGuid -and $entraById.ContainsKey($i.AzureADDeviceId)) {
            $e = $entraById[$i.AzureADDeviceId]
            [void]$matchedEntra.Add($e.DeviceId)
        }
        $eStale = $e -and $e.IsStale
        if (-not $i.IsStale -and -not $eStale) { continue }

        $combined.Add([pscustomobject]@{
            Name              = $i.DeviceName
            OperatingSystem   = $i.OperatingSystem
            Ownership         = $i.Ownership
            IntuneLastSync    = $i.IntuneLastSync
            EntraLastSignIn   = if ($e) { $e.EntraLastSignIn } else { $null }
            EntraEnabled      = if ($e) { $e.AccountEnabled } else { $null }
            TrustType         = if ($e) { $e.TrustType } else { $null }
            SerialNumber      = $i.SerialNumber
            PrimaryUserUPN    = $i.PrimaryUserUPN
            StaleIn           = if ($i.IsStale -and $eStale) { 'Both' } elseif ($i.IsStale) { 'Intune' } else { 'Entra' }
            Sources           = if ($e) { 'Intune+Entra' } else { 'IntuneOnly' }
            IsAutopilotObject = if ($e) { $e.IsAutopilotObject } else { $false }
            AzureADDeviceId   = $i.AzureADDeviceId
            IntuneDeviceId    = $i.IntuneDeviceId
            EntraObjectId     = if ($e) { $e.ObjectId } else { $null }
        })
    }
    foreach ($e in $entraRows) {
        if (-not $e.IsStale -or ($e.DeviceId -and $matchedEntra.Contains($e.DeviceId))) { continue }
        $combined.Add([pscustomobject]@{
            Name              = $e.DisplayName
            OperatingSystem   = $e.OperatingSystem
            Ownership         = $null
            IntuneLastSync    = $null
            EntraLastSignIn   = $e.EntraLastSignIn
            EntraEnabled      = $e.AccountEnabled
            TrustType         = $e.TrustType
            SerialNumber      = $null
            PrimaryUserUPN    = $null
            StaleIn           = 'Entra'
            Sources           = 'EntraOnly'
            IsAutopilotObject = $e.IsAutopilotObject
            AzureADDeviceId   = $e.DeviceId
            IntuneDeviceId    = $null
            EntraObjectId     = $e.ObjectId
        })
    }

    # 5. Export --------------------------------------------------------------------------
    $staleIntune = @($intuneRows | Where-Object IsStale | Select-Object * -ExcludeProperty IsStale | Sort-Object IntuneLastSync)
    $staleEntra  = @($entraRows  | Where-Object IsStale | Select-Object * -ExcludeProperty IsStale | Sort-Object EntraLastSignIn)
    $null = Export-Report -Data $staleIntune -Name "StaleDevices_Intune_${Days}d"
    $null = Export-Report -Data $staleEntra  -Name "StaleDevices_Entra_${Days}d"
    $null = Export-Report -Data ($combined | Sort-Object StaleIn, Name) -Name "StaleDevices_Combined_${Days}d"

    Write-Host ''
    Write-Log 'Summary:'
    Write-Log ("  Intune : {0} of {1} stale" -f $staleIntune.Count, $intuneRows.Count)
    Write-Log ("  Entra  : {0} of {1} stale ({2} of which are Autopilot objects)" -f `
        $staleEntra.Count, $entraRows.Count, @($staleEntra | Where-Object IsAutopilotObject).Count)
    $combined | Group-Object StaleIn | Sort-Object Name | ForEach-Object {
        Write-Log ('  Combined StaleIn={0,-6} {1}' -f $_.Name, $_.Count)
    }
    if ($combined.Count -gt 0) { $exitCode = 1 }
}
catch {
    if (-not $script:LogFile) { try { Initialize-Output -ScriptName 'StaleDevices' } catch { } }
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
exit $exitCode

#endregion
