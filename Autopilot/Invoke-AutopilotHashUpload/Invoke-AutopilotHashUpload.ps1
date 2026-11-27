#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Collects the Autopilot hardware hash and uploads it to Intune through Microsoft Graph,
    authenticating as an app registration with a certificate. Sets the group tag directly.

.DESCRIPTION
    Built to run in OOBE (Shift+F10 -> powershell) or on an existing device, without any
    PowerShell modules (no Microsoft.Graph, no WindowsAutopilotIntune, no MSAL):

      1. Reads the hardware hash (MDM_DevDetail_Ext01 in root/cimv2/mdm/dmmap), the serial
         number (Win32_BIOS) and, if present, the OEM product key.
      2. Optionally saves them to a CSV in the Intune import format (-OutputCsv).
      3. Builds a JWT client assertion signed with the certificate's private key (RS256,
         x5t header) and gets an app-only token from login.microsoftonline.com.
         Certificate source: -CertificateThumbprint (LocalMachine\My, then CurrentUser\My)
         or -PfxPath + -PfxPassword. Works with CNG and legacy CSP keys.
      4. Checks whether the serial number already exists in Autopilot. If it does, it reports
         it and, with -UpdateExisting, sets the group tag (and assigned user) through the
         updateDeviceProperties action instead of importing again.
      5. Imports the device (POST importedWindowsAutopilotDeviceIdentities) with groupTag,
         polls the import until 'complete' or 'error', and removes the import record.
      6. Unless -NoWait: polls windowsAutopilotDeviceIdentities until the device appears
         and reports the profile assignment status. With -WaitForProfile it keeps waiting
         until a deployment profile is assigned.

    App registration requirements:
      - API permission (Application): DeviceManagementServiceConfig.ReadWrite.All, admin consent.
      - The certificate's public key (.cer) uploaded under "Certificates & secrets".

.PARAMETER TenantId
    Tenant ID (GUID) or primary domain, e.g. 'contoso.onmicrosoft.com'.

.PARAMETER ClientId
    Application (client) ID of the app registration.

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate with private key in LocalMachine\My or CurrentUser\My.

.PARAMETER PfxPath
    Path to a .pfx file with the certificate and private key (e.g. on a USB stick).

.PARAMETER PfxPassword
    Password of the .pfx as SecureString.

.PARAMETER GroupTag
    Autopilot group tag (OrderID) to set.

.PARAMETER AssignedUser
    Optional UPN to assign to the device.

.PARAMETER UpdateExisting
    If the serial already exists in Autopilot: update its group tag/assigned user.

.PARAMETER OutputCsv
    Also save the hash to this CSV (Intune import format).

.PARAMETER TimeoutMinutes
    Maximum total waiting time for import and appearance (default 25).

.PARAMETER NoWait
    Do not wait for the device to appear in Autopilot after a successful import.

.PARAMETER WaitForProfile
    Keep waiting until a deployment profile is assigned (within -TimeoutMinutes).

.EXAMPLE
    # OOBE, certificate from USB stick
    $pw = Read-Host -AsSecureString 'PFX password'
    .\Invoke-AutopilotHashUpload.ps1 -TenantId contoso.onmicrosoft.com -ClientId <guid> `
        -PfxPath D:\autopilot.pfx -PfxPassword $pw -GroupTag 'NL-Standard' -WaitForProfile

.EXAMPLE
    .\Invoke-AutopilotHashUpload.ps1 -TenantId <guid> -ClientId <guid> -CertificateThumbprint <thumb> `
        -GroupTag 'NL-Kiosk' -UpdateExisting -OutputCsv C:\Temp\hash.csv

.NOTES
    Exit codes: 0 = imported / already present (and updated), 1 = import error, timeout or
    device already registered without -UpdateExisting, 2 = fatal error.
    Log: %ProgramData%\AutopilotHashUpload\
    Treat a PFX with private key as a secret: anyone holding it can register devices
    in your tenant. Prefer a dedicated app registration with only this permission.
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Thumbprint')]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [guid]$ClientId,

    [Parameter(Mandatory, ParameterSetName = 'Thumbprint')]
    [ValidatePattern('^[0-9A-Fa-f\s]{40,}$')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory, ParameterSetName = 'Pfx')]
    [string]$PfxPath,

    [Parameter(ParameterSetName = 'Pfx')]
    [securestring]$PfxPassword,

    [string]$GroupTag,

    [string]$AssignedUser,

    [switch]$UpdateExisting,

    [string]$OutputCsv,

    [ValidateRange(1, 240)]
    [int]$TimeoutMinutes = 25,

    [switch]$NoWait,

    [switch]$WaitForProfile,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'AutopilotHashUpload')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$GraphBase = 'https://graph.microsoft.com/v1.0'

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
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false } catch { }
    }
}

function ConvertTo-Base64Url([byte[]]$Bytes) {
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-DeviceHash {
    $devDetail = Get-CimInstance -Namespace 'root/cimv2/mdm/dmmap' -ClassName 'MDM_DevDetail_Ext01' `
        -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"
    if (-not $devDetail -or -not $devDetail.DeviceHardwareData) {
        throw 'Hardware hash not available (MDM_DevDetail_Ext01 empty). Run as administrator/SYSTEM.'
    }
    $serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
    if ($serial) { $serial = $serial.Trim() }

    $productKey = ''
    try {
        $sls = Get-CimInstance -ClassName SoftwareLicensingService
        if ($sls.OA3xOriginalProductKey) { $productKey = $sls.OA3xOriginalProductKey }
    }
    catch { }

    [pscustomobject]@{
        SerialNumber = $serial
        ProductKey   = $productKey
        HardwareHash = $devDetail.DeviceHardwareData
    }
}

function Get-AuthCertificate {
    if ($PSCmdlet.ParameterSetName -eq 'Pfx') {
        if (-not (Test-Path -LiteralPath $PfxPath)) { throw "PFX not found: $PfxPath" }
        $flags = [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
        $cert = New-Object Security.Cryptography.X509Certificates.X509Certificate2(
            (Resolve-Path -LiteralPath $PfxPath).ProviderPath, $PfxPassword, $flags)
    }
    else {
        $thumb = ($CertificateThumbprint -replace '\s', '').ToUpperInvariant()
        $cert = $null
        foreach ($store in 'Cert:\LocalMachine\My', 'Cert:\CurrentUser\My') {
            $c = Get-Item -LiteralPath (Join-Path $store $thumb) -ErrorAction SilentlyContinue
            if ($c) { $cert = $c; Write-Log "Certificate found in $store"; break }
        }
        if (-not $cert) { throw "Certificate $thumb not found in LocalMachine\My or CurrentUser\My." }
    }
    if (-not $cert.HasPrivateKey) { throw 'The certificate has no private key.' }
    if ($cert.NotAfter -lt (Get-Date)) { throw "Certificate expired on $($cert.NotAfter)." }
    Write-Log ("Certificate: {0} (thumbprint {1}, valid until {2:yyyy-MM-dd})" -f $cert.Subject, $cert.Thumbprint, $cert.NotAfter)
    return $cert
}

function New-SignatureRS256 {
    <# Signs data with SHA256/PKCS#1 using the certificate key; handles CNG and legacy CSP keys. #>
    param([Parameter(Mandatory)]$Cert, [Parameter(Mandatory)][byte[]]$Data)

    $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Cert)
    if (-not $rsa) { throw 'Could not access the RSA private key of the certificate.' }
    try {
        return $rsa.SignData($Data, [Security.Cryptography.HashAlgorithmName]::SHA256,
            [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    }
    catch {
        # Legacy CSP (PROV_RSA_FULL) keys cannot do SHA256: reopen the same key container
        # with the Enhanced RSA and AES provider (PROV_RSA_AES = 24).
        if ($rsa -isnot [Security.Cryptography.RSACryptoServiceProvider]) { throw }
        $info = $rsa.CspKeyContainerInfo
        $csp = New-Object Security.Cryptography.CspParameters(24, 'Microsoft Enhanced RSA and AES Cryptographic Provider', $info.KeyContainerName)
        $csp.KeyNumber = [int]$info.KeyNumber
        $csp.Flags = [Security.Cryptography.CspProviderFlags]::UseExistingKey
        if ($info.MachineKeyStore) { $csp.Flags = $csp.Flags -bor [Security.Cryptography.CspProviderFlags]::UseMachineKeyStore }
        $aes = New-Object Security.Cryptography.RSACryptoServiceProvider($csp)
        try { return $aes.SignData($Data, 'SHA256') } finally { $aes.Dispose() }
    }
}

function Get-GraphToken {
    param([Parameter(Mandatory)]$Cert)

    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    $header  = @{ alg = 'RS256'; typ = 'JWT'; x5t = (ConvertTo-Base64Url $Cert.GetCertHash()) } | ConvertTo-Json -Compress
    $payload = @{
        aud = $tokenUrl
        iss = "$ClientId"
        sub = "$ClientId"
        jti = [guid]::NewGuid().ToString()
        nbf = $now - 60
        iat = $now - 60
        exp = $now + 600
    } | ConvertTo-Json -Compress

    $unsigned = '{0}.{1}' -f (ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($header))),
                             (ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($payload)))
    $signature = New-SignatureRS256 -Cert $Cert -Data ([Text.Encoding]::UTF8.GetBytes($unsigned))
    $jwt = '{0}.{1}' -f $unsigned, (ConvertTo-Base64Url $signature)

    $body = @{
        client_id             = "$ClientId"
        scope                 = 'https://graph.microsoft.com/.default'
        grant_type            = 'client_credentials'
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = $jwt
    }
    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType 'application/x-www-form-urlencoded'
    }
    catch {
        $detail = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        throw "Token request failed: $detail"
    }
    Write-Log 'Access token obtained.'
    return $resp.access_token
}

function Invoke-Graph {
    <# Graph call with retry on 429/5xx (Retry-After honoured). #>
    param(
        [Parameter(Mandatory)][ValidateSet('GET', 'POST', 'DELETE', 'PATCH')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        $Body
    )
    $headers = @{ Authorization = "Bearer $script:Token" }
    $params = @{ Method = $Method; Uri = $Uri; Headers = $headers; ContentType = 'application/json' }
    if ($null -ne $Body) { $params.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10 -Compress)) }

    for ($attempt = 1; ; $attempt++) {
        try { return Invoke-RestMethod @params }
        catch {
            $status = $null
            if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response) {
                $status = [int]$_.Exception.Response.StatusCode
            }
            $detail = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
            if (($status -eq 429 -or $status -ge 500) -and $attempt -lt 5) {
                $wait = 5 * $attempt
                try { $ra = $_.Exception.Response.Headers['Retry-After']; if ($ra) { $wait = [int]$ra } } catch { }
                Write-Log "  Graph $status, retrying in $wait s..." 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }
            throw "Graph $Method $Uri failed ($status): $detail"
        }
    }
}

function Find-AutopilotDevice([string]$Serial) {
    $s = $Serial.Replace("'", "''")
    $uri = "$GraphBase/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$s')"
    $res = Invoke-Graph -Method GET -Uri $uri
    return @($res.value | Where-Object { $_.serialNumber -eq $Serial })
}

function Write-DeviceInfo($Device) {
    $get = { param($n) if ($Device.PSObject.Properties[$n]) { $Device.$n } else { '' } }
    Write-Log ("  Autopilot device: id {0}, serial {1}, group tag '{2}', profile status {3}, enrollment {4}" -f `
        (& $get 'id'), (& $get 'serialNumber'), (& $get 'groupTag'),
        (& $get 'deploymentProfileAssignmentStatus'), (& $get 'enrollmentState'))
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('HashUpload_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'HashUpload_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    Write-Log "=== Autopilot hash upload from $env:COMPUTERNAME (group tag '$GroupTag') ==="
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    # 1. Hash ---------------------------------------------------------------------------
    $hw = Get-DeviceHash
    if (-not $hw.SerialNumber) { throw 'No serial number found in Win32_BIOS.' }
    Write-Log ("Serial {0}, hash length {1}, product key {2}" -f $hw.SerialNumber, $hw.HardwareHash.Length,
        $(if ($hw.ProductKey) { 'present' } else { 'none' }))

    if ($OutputCsv) {
        $row = [pscustomobject][ordered]@{
            'Device Serial Number' = $hw.SerialNumber
            'Windows Product ID'   = $hw.ProductKey
            'Hardware Hash'        = $hw.HardwareHash
            'Group Tag'            = $GroupTag
            'Assigned User'        = $AssignedUser
        }
        $dir = Split-Path -Path $OutputCsv -Parent
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force -WhatIf:$false | Out-Null }
        # Intune's CSV import expects no quotes
        $row | ConvertTo-Csv -NoTypeInformation | ForEach-Object { $_ -replace '"', '' } |
            Set-Content -LiteralPath $OutputCsv -Encoding ASCII -WhatIf:$false
        Write-Log "Hash saved to $OutputCsv"
    }

    # 2. Authenticate -------------------------------------------------------------------
    $cert = Get-AuthCertificate
    $script:Token = Get-GraphToken -Cert $cert

    # 3. Duplicate check ----------------------------------------------------------------
    $existing = @(Find-AutopilotDevice $hw.SerialNumber)
    if ($existing.Count -gt 0) {
        Write-Log "Serial $($hw.SerialNumber) is already registered in Autopilot." 'WARN'
        foreach ($d in $existing) { Write-DeviceInfo $d }

        if (-not $UpdateExisting) {
            Write-Output "ALREADY REGISTERED: $($hw.SerialNumber) (use -UpdateExisting to change the group tag)"
            exit 1
        }

        $props = @{}
        if ($PSBoundParameters.ContainsKey('GroupTag')) { $props.groupTag = $GroupTag }
        if ($AssignedUser) { $props.userPrincipalName = $AssignedUser }
        if ($props.Count -eq 0) { Write-Output "ALREADY REGISTERED: $($hw.SerialNumber), nothing to update"; exit 0 }

        foreach ($d in $existing) {
            if ($PSCmdlet.ShouldProcess("Autopilot device $($d.id)", "updateDeviceProperties ($($props.Keys -join ', '))")) {
                Invoke-Graph -Method POST -Uri "$GraphBase/deviceManagement/windowsAutopilotDeviceIdentities/$($d.id)/updateDeviceProperties" -Body $props | Out-Null
                Write-Log "  Updated device $($d.id)."
            }
        }
        Write-Output "UPDATED: $($hw.SerialNumber) group tag '$GroupTag'"
        exit 0
    }

    # 4. Import -------------------------------------------------------------------------
    $import = @{
        '@odata.type'      = '#microsoft.graph.importedWindowsAutopilotDeviceIdentity'
        serialNumber       = $hw.SerialNumber
        productKey         = $hw.ProductKey
        hardwareIdentifier = $hw.HardwareHash
        groupTag           = $(if ($GroupTag) { $GroupTag } else { '' })
        assignedUserPrincipalName = $(if ($AssignedUser) { $AssignedUser } else { '' })
        state = @{
            '@odata.type'        = 'microsoft.graph.importedWindowsAutopilotDeviceIdentityState'
            deviceImportStatus   = 'pending'
            deviceRegistrationId = ''
            deviceErrorCode      = 0
            deviceErrorName      = ''
        }
    }

    if (-not $PSCmdlet.ShouldProcess("serial $($hw.SerialNumber)", "Import into Autopilot with group tag '$GroupTag'")) {
        Write-Output "WHATIF: would import $($hw.SerialNumber) with group tag '$GroupTag'"
        exit 0
    }

    $created = Invoke-Graph -Method POST -Uri "$GraphBase/deviceManagement/importedWindowsAutopilotDeviceIdentities" -Body $import
    Write-Log "Import started (id $($created.id)). Waiting for processing..."

    $status = 'pending'; $imp = $created
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 15
        $imp = Invoke-Graph -Method GET -Uri "$GraphBase/deviceManagement/importedWindowsAutopilotDeviceIdentities/$($created.id)"
        $status = $imp.state.deviceImportStatus
        Write-Log "  Import status: $status"
        if ($status -in 'complete', 'error') { break }
    }

    # Clean up the import record (the Autopilot device itself stays)
    try { Invoke-Graph -Method DELETE -Uri "$GraphBase/deviceManagement/importedWindowsAutopilotDeviceIdentities/$($created.id)" | Out-Null }
    catch { Write-Log "  Could not remove import record: $($_.Exception.Message)" 'WARN' }

    if ($status -eq 'error') {
        $err = '{0} ({1})' -f $imp.state.deviceErrorName, $imp.state.deviceErrorCode
        Write-Log "Import failed: $err" 'ERROR'
        Write-Output "FAILED: import of $($hw.SerialNumber): $err"
        exit 1
    }
    if ($status -ne 'complete') {
        Write-Log "Import not finished within $TimeoutMinutes minutes (status $status)." 'ERROR'
        Write-Output "TIMEOUT: import of $($hw.SerialNumber) still '$status'"
        exit 1
    }
    Write-Log 'Import complete.'

    # 5. Wait for the Autopilot device (and profile) ------------------------------------
    if ($NoWait) {
        Write-Output "IMPORTED: $($hw.SerialNumber) group tag '$GroupTag'"
        exit 0
    }

    Write-Log 'Waiting for the device to appear in Autopilot devices...'
    $device = $null
    while ((Get-Date) -lt $deadline) {
        $device = Find-AutopilotDevice $hw.SerialNumber | Select-Object -First 1
        if ($device) {
            $profileStatus = $device.deploymentProfileAssignmentStatus
            if (-not $WaitForProfile -or $profileStatus -in 'assignedInSync', 'assignedOutOfSync', 'assignedUnkownSyncState') { break }
            Write-Log "  Device present, profile status: $profileStatus"
        }
        Start-Sleep -Seconds 30
    }

    if (-not $device) {
        Write-Log 'Device did not appear within the timeout (import itself succeeded; it can take longer).' 'WARN'
        Write-Output "IMPORTED: $($hw.SerialNumber) (not yet visible in Autopilot devices)"
        exit 0
    }

    Write-DeviceInfo $device
    $profileStatus = $device.deploymentProfileAssignmentStatus
    if ($WaitForProfile -and $profileStatus -notlike 'assigned*') {
        Write-Log "No profile assigned within the timeout (status $profileStatus). Check the dynamic group on the group tag." 'WARN'
        Write-Output "IMPORTED, NO PROFILE YET: $($hw.SerialNumber) group tag '$($device.groupTag)' ($profileStatus)"
        exit 1
    }
    Write-Output "IMPORTED: $($hw.SerialNumber) group tag '$($device.groupTag)', profile status $profileStatus"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
