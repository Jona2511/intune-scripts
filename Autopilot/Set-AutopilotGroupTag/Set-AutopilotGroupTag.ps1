#Requires -Version 5.1

<#
.SYNOPSIS
    Bulk-changes Autopilot group tags from a CSV of serial numbers.

.DESCRIPTION
    Reads a CSV with the column SerialNumber and optionally GroupTag. With -GroupTag the
    same tag is applied to every serial number (overrides a GroupTag column).
    The input delimiter (';', ',' or tab) is detected automatically.

    For each serial number the script looks up the windowsAutopilotDeviceIdentity:
      - up to -LookupThreshold serials: one filtered query per serial
        (contains(serialNumber,'...'), then an exact match);
      - more serials: all Autopilot devices are fetched once and indexed by serial.

    The tag is changed with POST .../windowsAutopilotDeviceIdentities/{id}/updateDeviceProperties.
    Supports -WhatIf.

    Result per serial (CSV): Updated, Unchanged, NotFound, Duplicate, Failed, Skipped,
    or WouldUpdate (under -WhatIf).
      Duplicate = serial found more than once in Autopilot, or listed in the CSV more
                  than once with different tags. Such serials are never changed.
      Skipped   = empty group tag in the CSV.

.PARAMETER TenantId
    Tenant ID or tenant domain (e.g. contoso.onmicrosoft.com).

.PARAMETER CsvPath
    CSV with the column SerialNumber and optionally GroupTag.

.PARAMETER GroupTag
    One group tag for every serial in the CSV. Overrides the GroupTag column.

.PARAMETER ClientId
    App registration (client) ID for app-only sign-in with a certificate. Without
    -ClientId the script signs in interactively (delegated).

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate (CurrentUser\My or LocalMachine\My) for -ClientId.

.PARAMETER OutputDirectory
    Folder for the result CSV and the log. Default: current folder.

.PARAMETER Delimiter
    Delimiter of the result CSV. Default ';' (Dutch Excel).

.PARAMETER LookupThreshold
    Up to this number of serials a filtered query per serial is used; above it, all
    Autopilot devices are fetched once.

.EXAMPLE
    .\Set-AutopilotGroupTag.ps1 -TenantId contoso.onmicrosoft.com -CsvPath .\tags.csv -WhatIf

.EXAMPLE
    .\Set-AutopilotGroupTag.ps1 -TenantId $tid -ClientId $app -CertificateThumbprint $thumb `
        -CsvPath .\serials.csv -GroupTag 'NL-Office'

.NOTES
    Graph permissions
      Application : DeviceManagementServiceConfig.ReadWrite.All
                    Organization.Read.All (optional, for the tenant domain in file names)
      Delegated   : DeviceManagementServiceConfig.ReadWrite.All, User.Read
                    (role: Intune Administrator or equivalent)

    Dependency: Microsoft.Graph.Authentication.
    The new tag shows in Intune within minutes. The [OrderID] on the Entra device, and
    therefore dynamic group membership, follows later.

    Exit codes: 0 = all Updated/Unchanged, 1 = NotFound/Duplicate/Failed/Skipped present,
    2 = fatal error.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$CsvPath,
    [string]$GroupTag,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$OutputDirectory = (Get-Location).Path,
    [string]$Delimiter = ';',
    [ValidateRange(1, 1000)][int]$LookupThreshold = 25
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
    # 1. Read the input CSV (before connecting, so input errors fail fast) ----------------
    if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV not found: $CsvPath" }
    $firstLine = Get-Content -LiteralPath $CsvPath -TotalCount 1 -Encoding UTF8
    $inDelim = if ($firstLine -match ';') { ';' } elseif ($firstLine -match "`t") { "`t" } else { ',' }
    $csv = @(Import-Csv -LiteralPath $CsvPath -Delimiter $inDelim -Encoding UTF8)
    if ($csv.Count -eq 0) { throw 'The CSV contains no rows.' }
    if (-not $csv[0].PSObject.Properties['SerialNumber']) { throw "The CSV has no 'SerialNumber' column." }
    $hasTagColumn = [bool]$csv[0].PSObject.Properties['GroupTag']
    if (-not $hasTagColumn -and -not $PSBoundParameters.ContainsKey('GroupTag')) {
        throw "The CSV has no 'GroupTag' column; specify -GroupTag."
    }

    $inputRows = foreach ($row in $csv) {
        $serial = "$(Get-Value $row 'SerialNumber')".Trim()
        if (-not $serial) { continue }
        $tag = if ($PSBoundParameters.ContainsKey('GroupTag')) { $GroupTag } else { "$(Get-Value $row 'GroupTag')".Trim() }
        [pscustomobject]@{ SerialNumber = $serial; GroupTag = $tag }
    }
    $inputRows = @($inputRows)

    # 2. Connect -------------------------------------------------------------------------
    Connect-Tenant -Scopes 'DeviceManagementServiceConfig.ReadWrite.All', 'User.Read'
    $script:TenantLabel = Get-TenantLabel
    Initialize-Output -ScriptName 'AutopilotGroupTag'
    Write-Log "=== Autopilot group tag update for $script:TenantLabel ($(if ($WhatIfPreference) { 'WhatIf' } else { 'Apply' })) ==="
    Write-Log "Input: $($inputRows.Count) row(s) from $CsvPath (delimiter '$inDelim')"
    if ($PSBoundParameters.ContainsKey('GroupTag') -and $hasTagColumn) {
        Write-Log "-GroupTag '$GroupTag' overrides the GroupTag column in the CSV." 'WARN'
    }

    # 3. Resolve CSV duplicates ------------------------------------------------------------
    $results = New-Object System.Collections.Generic.List[object]
    $work    = New-Object System.Collections.Generic.List[object]
    foreach ($grp in ($inputRows | Group-Object -Property SerialNumber)) {
        $tags = @($grp.Group | ForEach-Object { $_.GroupTag } | Select-Object -Unique)
        if ($tags.Count -gt 1) {
            $results.Add([pscustomobject]@{
                SerialNumber = $grp.Name; NewGroupTag = ($tags -join ' | '); OldGroupTag = $null
                Status = 'Duplicate'; Detail = 'Serial appears in the CSV with different group tags'
                AutopilotId = $null; Model = $null
            })
        }
        else { $work.Add($grp.Group[0]) }
    }

    # 4. Look up Autopilot devices --------------------------------------------------------
    $base  = 'v1.0/deviceManagement/windowsAutopilotDeviceIdentities'
    $index = @{}   # serial (case-insensitive) -> list of identities
    if ($work.Count -le $LookupThreshold) {
        Write-Log "Looking up $($work.Count) serial(s) individually"
        foreach ($w in $work) {
            $filter = [uri]::EscapeDataString("contains(serialNumber,'$($w.SerialNumber.Replace("'", "''"))')")
            $found  = Get-GraphAll -Uri "$base`?`$filter=$filter"
            $index[$w.SerialNumber] = @($found | Where-Object { "$(Get-Value $_ 'serialNumber')".Trim() -eq $w.SerialNumber })
        }
    }
    else {
        Write-Log 'Fetching all Autopilot devices once'
        $all = Get-GraphAll -Uri $base
        Write-Log "  $($all.Count) Autopilot device(s) in the tenant"
        foreach ($d in $all) {
            $s = "$(Get-Value $d 'serialNumber')".Trim()
            if (-not $s) { continue }
            if (-not $index.ContainsKey($s)) { $index[$s] = New-Object System.Collections.Generic.List[object] }
            $index[$s].Add($d)
        }
    }

    # 5. Update ------------------------------------------------------------------------------
    foreach ($w in $work) {
        $r = [pscustomobject]@{
            SerialNumber = $w.SerialNumber; NewGroupTag = $w.GroupTag; OldGroupTag = $null
            Status = $null; Detail = $null; AutopilotId = $null; Model = $null
        }
        $hits = if ($index.ContainsKey($w.SerialNumber)) { @($index[$w.SerialNumber]) } else { @() }

        if ($hits.Count -eq 0) {
            $r.Status = 'NotFound'; $r.Detail = 'Serial not found in Autopilot'
        }
        elseif ($hits.Count -gt 1) {
            $r.Status = 'Duplicate'
            $r.Detail = 'Serial registered multiple times: ' + (($hits | ForEach-Object { Get-Value $_ 'id' }) -join ', ')
        }
        elseif (-not $w.GroupTag) {
            $dev = $hits[0]
            $r.AutopilotId = Get-Value $dev 'id'; $r.Model = Get-Value $dev 'model'; $r.OldGroupTag = Get-Value $dev 'groupTag'
            $r.Status = 'Skipped'; $r.Detail = 'Empty group tag in the CSV'
        }
        else {
            $dev = $hits[0]
            $r.AutopilotId = Get-Value $dev 'id'
            $r.Model       = Get-Value $dev 'model'
            $r.OldGroupTag = "$(Get-Value $dev 'groupTag')"

            if ($r.OldGroupTag -ceq $w.GroupTag) {
                $r.Status = 'Unchanged'
            }
            elseif ($PSCmdlet.ShouldProcess("$($w.SerialNumber) ('$($r.OldGroupTag)' -> '$($w.GroupTag)')", 'Set Autopilot group tag')) {
                try {
                    $null = Invoke-GraphApi -Method POST -Uri "$base/$($r.AutopilotId)/updateDeviceProperties" `
                        -Body @{ groupTag = $w.GroupTag }
                    $r.Status = 'Updated'
                }
                catch {
                    $r.Status = 'Failed'; $r.Detail = $_.Exception.Message
                }
            }
            else { $r.Status = 'WouldUpdate' }
        }

        $lvl = if ($r.Status -in 'Updated', 'Unchanged', 'WouldUpdate') { 'INFO' } else { 'WARN' }
        Write-Log ("  {0,-20} {1,-11} '{2}' -> '{3}' {4}" -f $r.SerialNumber, $r.Status, $r.OldGroupTag, $r.NewGroupTag, $r.Detail) $lvl
        $results.Add($r)
    }

    # 6. Report ----------------------------------------------------------------------------
    $null = Export-Report -Data $results.ToArray() -Name 'AutopilotGroupTag'

    Write-Host ''
    Write-Host 'Summary:'
    $results | Group-Object Status | Sort-Object Name | ForEach-Object {
        Write-Log ('  {0,-11} {1}' -f $_.Name, $_.Count)
    }
    if (@($results | Where-Object { $_.Status -in 'NotFound', 'Duplicate', 'Failed', 'Skipped' }).Count -gt 0) {
        $exitCode = 1
    }
}
catch {
    if (-not $script:LogFile) { try { Initialize-Output -ScriptName 'AutopilotGroupTag' } catch { } }
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
exit $exitCode

#endregion
