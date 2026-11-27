#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Detects unexpected members of the local Administrators group and reports or removes them.

.DESCRIPTION
    Reads the members of the local Administrators group (found by SID S-1-5-32-544, so it
    also works on Dutch Windows, where the group is called "Administrators"). Uses ADSI
    instead of Get-LocalGroupMember, because that cmdlet fails on Entra ID members and
    orphaned SIDs.

    Each member is compared against an allow list. Entries can be:
      - account names:  'CONTOSO\IT-Admins', 'LAPSAdmin', '*\Domain Admins'
      - SIDs:           'S-1-12-1-1234-...'
      - SID wildcards:  'S-1-5-21-*-512'
    Always allowed (built in):
      - the built-in local Administrator account (local SID ending in -500)
      - Domain Admins (domain SID ending in -512) on domain-joined devices

    Entra ID role and group members (e.g. Global Administrator, Microsoft Entra Joined
    Device Local Administrator) appear as unresolvable S-1-12-1-... SIDs. Allow them with
    -AllowedEntraObjectIds <object ID of the role/group>; the script converts the GUID to the SID.

    Default mode is report only. With -Remove, unexpected members are removed (supports -WhatIf).
    Safety: no removal happens if afterwards no enabled administrator would remain.

.PARAMETER AllowedMembers
    Allowed names/SIDs (wildcards allowed).

.PARAMETER AllowListPath
    Text file with extra allowed entries, one per line ('#' = comment).

.PARAMETER AllowedEntraObjectIds
    Object IDs (GUIDs) of Entra ID roles or groups that may be members.

.PARAMETER Remove
    Removes unexpected members. Without this switch the script only reports.

.EXAMPLE
    .\Invoke-LocalAdminCheck.ps1 -AllowedMembers 'CONTOSO\Workstation Admins'

.EXAMPLE
    .\Invoke-LocalAdminCheck.ps1 -AllowedEntraObjectIds '9a1b...','4c2d...' -Remove -WhatIf

.NOTES
    Exit codes: 0 = no unexpected members (afterwards), 1 = unexpected members present /
    removal failed or blocked, 2 = fatal error.
    Log: %ProgramData%\LocalAdminCheck\
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]]$AllowedMembers = @(),

    [string]$AllowListPath,

    [guid[]]$AllowedEntraObjectIds = @(),

    [switch]$Remove,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'LocalAdminCheck')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

function ConvertFrom-EntraObjectId {
    <# Converts an Entra ID object ID (GUID) to the S-1-12-1-... SID used on the device. #>
    param([Parameter(Mandatory)][guid]$ObjectId)
    $bytes = $ObjectId.ToByteArray()
    $parts = 0..3 | ForEach-Object { [BitConverter]::ToUInt32($bytes, $_ * 4) }
    'S-1-12-1-' + ($parts -join '-')
}

function Get-LocalMachineSid {
    $u = Get-LocalUser | Select-Object -First 1
    if ($u) { return $u.SID.AccountDomainSid.Value }
    return $null
}

function Get-AdminGroupMembers {
    param([Parameter(Mandatory)]$Group)

    foreach ($m in @($Group.psbase.Invoke('Members'))) {
        $type   = $m.GetType()
        $path   = $type.InvokeMember('ADsPath', 'GetProperty', $null, $m, $null)
        $class  = $type.InvokeMember('Class', 'GetProperty', $null, $m, $null)
        $sidRaw = $type.InvokeMember('objectSid', 'GetProperty', $null, $m, $null)
        $sid    = (New-Object Security.Principal.SecurityIdentifier($sidRaw, 0)).Value

        # WinNT://DOMAIN/name, WinNT://DOMAIN/COMPUTER/name (local) or WinNT://S-1-... (unresolved)
        $parts = $path -replace '^WinNT://', '' -split '/'
        $name = switch ($parts.Count) {
            3       { '{0}\{1}' -f $parts[1], $parts[2] }
            2       { '{0}\{1}' -f $parts[0], $parts[1] }
            default { $null }
        }
        if (-not $name) {
            try { $name = (New-Object Security.Principal.SecurityIdentifier($sid)).Translate([Security.Principal.NTAccount]).Value }
            catch { $name = '(unresolved)' }
        }

        [pscustomobject]@{
            Name    = $name
            Sid     = $sid
            Class   = $class
            AdsPath = $path
            IsLocal = ($parts.Count -eq 3)
        }
    }
}

function Test-Allowed {
    param($Member, [string[]]$AllowList)
    $shortName = ($Member.Name -split '\\')[-1]
    foreach ($entry in $AllowList) {
        if ($Member.Sid -like $entry -or $Member.Name -like $entry -or $shortName -like $entry) {
            return $entry
        }
    }
    return $null
}

function Test-EnabledAdmin {
    <# True if the member is an enabled local user, or a domain/Entra account or group (assumed usable). #>
    param($Member)
    if ($Member.IsLocal -and $Member.Class -eq 'User') {
        $u = Get-LocalUser -SID $Member.Sid -ErrorAction SilentlyContinue
        return [bool]($u -and $u.Enabled)
    }
    return $true
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('LocalAdmins_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'LocalAdmins_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    $mode = if ($Remove) { if ($WhatIfPreference) { 'Remove (WhatIf)' } else { 'Remove' } } else { 'Report' }
    Write-Log "=== Local admin check on $env:COMPUTERNAME ($mode) ==="

    # 1. Build the allow list -------------------------------------------------------------
    $allow = New-Object System.Collections.Generic.List[string]
    $machineSid = Get-LocalMachineSid
    if ($machineSid) { $allow.Add("$machineSid-500") }                      # built-in Administrator
    if ((Get-CimInstance Win32_ComputerSystem).PartOfDomain) {
        $allow.Add('S-1-5-21-*-512')                                          # Domain Admins
    }
    foreach ($a in $AllowedMembers) { if ($a.Trim()) { $allow.Add($a.Trim()) } }
    foreach ($id in $AllowedEntraObjectIds) { $allow.Add((ConvertFrom-EntraObjectId $id)) }
    if ($AllowListPath) {
        if (-not (Test-Path -LiteralPath $AllowListPath)) { throw "Allow list not found: $AllowListPath" }
        Get-Content -LiteralPath $AllowListPath -Encoding UTF8 |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') } |
            ForEach-Object { $allow.Add($_) }
    }
    Write-Log ("Allow list ({0}): {1}" -f $allow.Count, ($allow -join '; '))

    # 2. Read the group -------------------------------------------------------------------
    $groupName = (New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate(
        [Security.Principal.NTAccount]).Value.Split('\')[-1]
    $group = [ADSI]"WinNT://$env:COMPUTERNAME/$groupName,group"
    $members = @(Get-AdminGroupMembers -Group $group)
    Write-Log "Group '$groupName' has $($members.Count) member(s)."

    $unexpected = New-Object System.Collections.Generic.List[object]
    foreach ($m in $members) {
        $match = Test-Allowed -Member $m -AllowList $allow
        if ($match) {
            Write-Log ("  OK          {0} [{1}] {2}  (rule: {3})" -f $m.Name, $m.Class, $m.Sid, $match)
        }
        else {
            Write-Log ("  UNEXPECTED  {0} [{1}] {2}" -f $m.Name, $m.Class, $m.Sid) 'WARN'
            $unexpected.Add($m)
        }
    }

    if ($unexpected.Count -eq 0) {
        Write-Output "COMPLIANT: $($members.Count) member(s), all allowed"
        exit 0
    }

    $summary = ($unexpected | ForEach-Object { $_.Name }) -join '; '
    if (-not $Remove) {
        Write-Output "NON-COMPLIANT: unexpected local admin(s): $summary"
        exit 1
    }

    # 3. Safety check: an enabled admin must remain ----------------------------------------
    $remaining = @($members | Where-Object { $unexpected -notcontains $_ })
    $usable    = @($remaining | Where-Object { Test-EnabledAdmin $_ })
    if ($usable.Count -eq 0) {
        Write-Log 'Removal blocked: no enabled administrator would remain. Extend the allow list.' 'ERROR'
        Write-Output "NON-COMPLIANT: removal blocked (no admin would remain): $summary"
        exit 1
    }

    # 4. Remove ---------------------------------------------------------------------------
    $failed = New-Object System.Collections.Generic.List[string]
    foreach ($m in $unexpected) {
        if ($PSCmdlet.ShouldProcess("$($m.Name) ($($m.Sid))", "Remove from $groupName")) {
            try {
                $group.psbase.Invoke('Remove', $m.AdsPath)
                Write-Log "  Removed: $($m.Name) ($($m.Sid))"
            }
            catch {
                $err = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
                Write-Log "  Removing $($m.Name) failed: $err" 'ERROR'
                $failed.Add($m.Name)
            }
        }
        else { $failed.Add($m.Name) }   # WhatIf
    }

    if ($failed.Count -gt 0) {
        Write-Output ("NON-COMPLIANT: not removed: {0}" -f ($failed -join '; '))
        exit 1
    }
    Write-Log 'All unexpected members removed. Changes take effect at the user''s next sign-in.'
    Write-Output "REMEDIATED: removed $summary"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
