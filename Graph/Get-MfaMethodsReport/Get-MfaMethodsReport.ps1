#Requires -Version 5.1

<#
.SYNOPSIS
    Reports users without MFA and users whose only MFA methods are phone-based (SMS/voice).

.DESCRIPTION
    Source: /reports/authenticationMethods/userRegistrationDetails (v1.0). Microsoft refreshes
    this data periodically, so it is not real-time.

    Categories:
      NoMfa      isMfaRegistered = false
      PhoneOnly  MFA registered, but every MFA method is a phone method
                 (mobilePhone, alternateMobilePhone, officePhone, sms, voice).
                 email, securityQuestion and temporaryAccessPass do not count as MFA methods.

    Account status (enabled/disabled) and last sign-in come from /users. Last sign-in
    (signInActivity) needs Entra ID P1/P2 and AuditLog.Read.All; if it is not available the
    script continues without it.

    Output: Mfa_Findings_<domain>_<date>.csv (and with -ExportAll also Mfa_AllUsers_...).

.PARAMETER TenantId
    Tenant ID or initial domain.

.PARAMETER ClientId
    App registration (client) ID for app-only authentication. Without it: interactive (delegated).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication.

.PARAMETER ExcludeGuests
    Leave guest accounts out.

.PARAMETER ExcludeDisabled
    Leave disabled accounts out.

.PARAMETER SkipSignInActivity
    Do not request signInActivity (e.g. for tenants without P1).

.PARAMETER ExportAll
    Also export all users with their registration details.

.EXAMPLE
    .\Get-MfaMethodsReport.ps1 -TenantId contoso.onmicrosoft.com -ExcludeGuests -ExcludeDisabled

.NOTES
    Required Graph permissions (application or delegated):
      AuditLog.Read.All                  userRegistrationDetails and signInActivity
      UserAuthenticationMethod.Read.All  userRegistrationDetails
      User.Read.All                      accountEnabled, userType
      Organization.Read.All              tenant name (optional)
    Delegated: the user also needs e.g. Reports Reader, Security Reader or Global Reader.
    Licensing: the authentication methods registration report requires Entra ID P1/P2.

    Dependency: Microsoft.Graph.Authentication.
    Exit codes: 0 = no findings, 1 = findings, 2 = fatal error.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,

    [switch]$ExcludeGuests,
    [switch]$ExcludeDisabled,
    [switch]$SkipSignInActivity,
    [switch]$ExportAll,

    [string]$OutputDirectory = (Get-Location).Path,
    [string]$Delimiter = ';'
)

# StrictMode 1 only: Graph objects may omit properties, and level 2+ throws on that.
Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

$Scopes = @('AuditLog.Read.All', 'UserAuthenticationMethod.Read.All', 'User.Read.All', 'Organization.Read.All')

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

#region Main ---------------------------------------------------------------------------------

$PhoneMethods = @('mobilePhone', 'alternateMobilePhone', 'officePhone', 'sms', 'voice')
$NonMfaMethods = @('email', 'securityQuestion', 'temporaryAccessPass')

$exitCode = 2
try {
    Initialize-Output 'Mfa'
    Connect-Tenant -Scopes $Scopes
    $script:Tenant = Get-TenantLabel
    Write-Log "=== MFA / sign-in methods report for $($script:Tenant.Name) ($($script:Tenant.Domain)) ==="

    # 1. Registration details -------------------------------------------------------------------
    $details = @(Invoke-Graph -Uri 'v1.0/reports/authenticationMethods/userRegistrationDetails')
    Write-Log "Registration details: $($details.Count) user(s)"

    # 2. Account status + last sign-in ----------------------------------------------------------
    $users = $null
    $signIn = $false
    if (-not $SkipSignInActivity) {
        try {
            $users = @(Invoke-Graph -Uri 'v1.0/users?$select=id,accountEnabled,userType,signInActivity&$top=500')
            $signIn = $true
        }
        catch {
            Write-Log "signInActivity not available ($($_.Exception.Message)); continuing without last sign-in." 'WARN'
        }
    }
    if (-not $signIn) {
        $users = @(Invoke-Graph -Uri 'v1.0/users?$select=id,accountEnabled,userType&$top=999')
    }
    $userById = @{}
    foreach ($u in $users) { $userById[$u.id] = $u }

    # 3. Classify -------------------------------------------------------------------------------
    $all = New-Object System.Collections.Generic.List[object]
    foreach ($d in $details) {
        $u = $userById[$d.id]
        if (-not $u) { continue }                      # deleted in the meantime
        if ($ExcludeGuests -and $u.userType -eq 'Guest') { continue }
        if ($ExcludeDisabled -and -not $u.accountEnabled) { continue }

        $methods    = @($d.methodsRegistered | Where-Object { $_ })
        $mfaMethods = @($methods | Where-Object { $_ -notin $NonMfaMethods })
        $strong     = @($mfaMethods | Where-Object { $_ -notin $PhoneMethods })

        $finding = if (-not $d.isMfaRegistered -or $mfaMethods.Count -eq 0) { 'NoMfa' }
                   elseif ($strong.Count -eq 0) { 'PhoneOnly' }
                   else { '' }

        $lastSignIn = ''
        if ($signIn -and $u.signInActivity) {
            $dates = @(
                (ConvertTo-DateTimeUtc $u.signInActivity.lastSignInDateTime),
                (ConvertTo-DateTimeUtc $u.signInActivity.lastNonInteractiveSignInDateTime)
            ) | Where-Object { $_ } | Sort-Object -Descending
            if ($dates) { $lastSignIn = Format-LocalDate @($dates)[0] }
        }

        $all.Add([pscustomobject][ordered]@{
            Finding               = $finding
            UserPrincipalName     = $d.userPrincipalName
            DisplayName           = $d.userDisplayName
            AccountEnabled        = $u.accountEnabled
            UserType              = $u.userType
            IsAdmin               = $d.isAdmin
            IsMfaRegistered       = $d.isMfaRegistered
            IsMfaCapable          = $d.isMfaCapable
            IsPasswordlessCapable = $d.isPasswordlessCapable
            DefaultMfaMethod      = $d.defaultMfaMethod
            PreferredSecondary    = $d.userPreferredMethodForSecondaryAuthentication
            MethodsRegistered     = ($methods -join ', ')
            LastSignIn            = $lastSignIn
            ReportUpdated         = (Format-LocalDate $d.lastUpdatedDateTime)
        })
    }

    $findings = @($all | Where-Object { $_.Finding })
    $null = Export-Report ($findings | Sort-Object Finding, UserPrincipalName) 'Mfa_Findings'
    if ($ExportAll) { $null = Export-Report ($all | Sort-Object UserPrincipalName) 'Mfa_AllUsers' }

    # 4. Summary --------------------------------------------------------------------------------
    $noMfa     = @($findings | Where-Object { $_.Finding -eq 'NoMfa' })
    $phoneOnly = @($findings | Where-Object { $_.Finding -eq 'PhoneOnly' })
    $admins    = @($findings | Where-Object { $_.IsAdmin })

    Write-Log ("Summary: {0} user(s) checked, {1} without MFA, {2} phone/SMS only, {3} admin(s) among the findings." -f `
        $all.Count, $noMfa.Count, $phoneOnly.Count, $admins.Count)
    if ($admins.Count -gt 0) {
        Write-Log ("Admins with weak/no MFA: {0}" -f (($admins | ForEach-Object { "$($_.UserPrincipalName) [$($_.Finding)]" }) -join '; ')) 'WARN'
    }

    $exitCode = if ($findings.Count -gt 0) { 1 } else { 0 }
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
