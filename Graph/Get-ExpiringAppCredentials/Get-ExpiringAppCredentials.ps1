#Requires -Version 5.1

<#
.SYNOPSIS
    Finds client secrets and certificates of app registrations (and optionally enterprise apps)
    that expire within X days.

.DESCRIPTION
    Checks passwordCredentials (secrets) and keyCredentials (certificates) of:
      - app registrations (/applications)
      - with -IncludeServicePrincipals: enterprise apps (/servicePrincipals) of type
        'Application', which includes SAML signing certificates of gallery/SAML apps.

    Reports credentials whose end date is within -Days (default 30). Already expired
    credentials are included only with -IncludeExpired.

    Column HasNewerCredential = True means the same object already has another credential
    of the same type (and usage) that is valid beyond the window: the credential has probably
    been rotated already and only the old one still needs to be removed.

    Output: AppCredentials_<domain>_<date>.csv

.PARAMETER TenantId
    Tenant ID or initial domain.

.PARAMETER ClientId
    App registration (client) ID for app-only authentication. Without it: interactive (delegated).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication.

.PARAMETER Days
    Window in days (default 30).

.PARAMETER IncludeExpired
    Also report credentials that have already expired.

.PARAMETER IncludeServicePrincipals
    Also check enterprise apps (service principals), e.g. SAML signing certificates.

.EXAMPLE
    .\Get-ExpiringAppCredentials.ps1 -TenantId contoso.onmicrosoft.com -Days 60 -IncludeServicePrincipals

.NOTES
    Required Graph permissions (application or delegated):
      Application.Read.All    applications, servicePrincipals and their owners
      Organization.Read.All   tenant name (optional)
    Owner UPNs of user owners may additionally need User.ReadBasic.All in some tenants.

    Dependency: Microsoft.Graph.Authentication.
    Exit codes: 0 = nothing expiring, 1 = expiring/expired credentials found, 2 = fatal error.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,

    [ValidateRange(1, 3650)]
    [int]$Days = 30,

    [switch]$IncludeExpired,
    [switch]$IncludeServicePrincipals,

    [string]$OutputDirectory = (Get-Location).Path,
    [string]$Delimiter = ';'
)

# StrictMode 1 only: Graph objects may omit properties, and level 2+ throws on that.
Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

$Scopes = @('Application.Read.All', 'Organization.Read.All')

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

function Get-CredentialRows {
    param([Parameter(Mandatory)]$Object, [Parameter(Mandatory)][string]$ObjectType)

    $now   = (Get-Date).ToUniversalTime()
    $limit = $now.AddDays($Days)

    $owners = (@($Object.owners) | Where-Object { $_ } | ForEach-Object {
            if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName }
        }) -join '; '

    $creds = New-Object System.Collections.Generic.List[object]
    foreach ($c in $Object.passwordCredentials) {
        $creds.Add([pscustomobject]@{ Kind = 'Secret'; Usage = ''; Cred = $c; End = (ConvertTo-DateTimeUtc $c.endDateTime) })
    }
    foreach ($c in $Object.keyCredentials) {
        $creds.Add([pscustomobject]@{ Kind = 'Certificate'; Usage = "$($c.usage)"; Cred = $c; End = (ConvertTo-DateTimeUtc $c.endDateTime) })
    }

    foreach ($x in $creds) {
        if (-not $x.End -or $x.End -gt $limit) { continue }
        $expired = $x.End -lt $now
        if ($expired -and -not $IncludeExpired) { continue }

        $newer = @($creds | Where-Object {
                $_.Kind -eq $x.Kind -and $_.Usage -eq $x.Usage -and $_.End -and $_.End -gt $limit
            }).Count -gt 0

        $name = if ($x.Cred.displayName) { $x.Cred.displayName }
                elseif ($x.Kind -eq 'Secret' -and $x.Cred.hint) { "hint: $($x.Cred.hint)***" }
                else { '' }

        [pscustomobject][ordered]@{
            Status             = $(if ($expired) { 'Expired' } else { 'Expiring' })
            ObjectType         = $ObjectType
            AppName            = $Object.displayName
            AppId              = $Object.appId
            ObjectId           = $Object.id
            CredentialType     = $x.Kind
            Usage              = $x.Usage
            CredentialName     = $name
            KeyId              = $x.Cred.keyId
            Start              = (Format-LocalDate $x.Cred.startDateTime)
            End                = (Format-LocalDate $x.End)
            DaysLeft           = [int][math]::Floor(($x.End - $now).TotalDays)
            HasNewerCredential = $newer
            Owners             = $owners
        }
    }
}

#endregion

#region Main ---------------------------------------------------------------------------------

$exitCode = 2
try {
    Initialize-Output 'AppCredentials'
    Connect-Tenant -Scopes $Scopes
    $script:Tenant = Get-TenantLabel
    Write-Log "=== Expiring app credentials (<= $Days days) for $($script:Tenant.Name) ($($script:Tenant.Domain)) ==="

    $rows = New-Object System.Collections.Generic.List[object]

    $apps = @(Invoke-Graph -Uri 'v1.0/applications?$select=id,appId,displayName,passwordCredentials,keyCredentials&$expand=owners($select=id,displayName,userPrincipalName)')
    Write-Log "App registrations: $($apps.Count)"
    foreach ($a in $apps) { foreach ($r in (Get-CredentialRows -Object $a -ObjectType 'AppRegistration')) { $rows.Add($r) } }

    if ($IncludeServicePrincipals) {
        $sps = @(Invoke-Graph -Uri 'v1.0/servicePrincipals?$select=id,appId,displayName,servicePrincipalType,passwordCredentials,keyCredentials&$expand=owners($select=id,displayName,userPrincipalName)')
        $sps = @($sps | Where-Object {
                $_.servicePrincipalType -eq 'Application' -and ($_.passwordCredentials -or $_.keyCredentials)
            })
        Write-Log "Enterprise apps with credentials: $($sps.Count)"
        foreach ($s in $sps) { foreach ($r in (Get-CredentialRows -Object $s -ObjectType 'EnterpriseApp')) { $rows.Add($r) } }
    }

    $null = Export-Report ($rows | Sort-Object DaysLeft, AppName) 'AppCredentials'

    $expired  = @($rows | Where-Object { $_.Status -eq 'Expired' })
    $expiring = @($rows | Where-Object { $_.Status -eq 'Expiring' })
    $rotated  = @($rows | Where-Object { $_.HasNewerCredential })

    if ($rows.Count -gt 0) {
        Write-Host ''
        $rows | Sort-Object DaysLeft | Format-Table Status, DaysLeft, AppName, CredentialType, Usage, End, HasNewerCredential -AutoSize |
            Out-String | Write-Host
    }
    Write-Log ("Summary: {0} expiring within {1} days, {2} expired, {3} of them already have a newer credential." -f `
        $expiring.Count, $Days, $expired.Count, $rotated.Count)

    $exitCode = if ($rows.Count -gt 0) { 1 } else { 0 }
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
