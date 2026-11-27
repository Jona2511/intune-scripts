#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Checks BitLocker status and makes sure recovery keys are escrowed to Entra ID.

.DESCRIPTION
    For every fixed drive (optionally also removable drives) the script checks:

      - Whether the drive is encrypted. Unencrypted drives are reported, not encrypted
        by this script.
      - Whether protection is suspended.
      - Whether each RecoveryPassword protector has been backed up to Entra ID.

    Escrow detection: Windows logs event 845 in
    'Microsoft-Windows-BitLocker/BitLocker Management' when a recovery key is
    successfully backed up to Entra ID. A protector without such an event (for example
    because the event log has rolled over) is treated as not escrowed and is backed up
    again with BackupToAAD-BitLockerKeyProtector. A repeat backup is harmless.

    The recovery password itself is never written to the screen or the log.

    Works as an Intune remediation:
      detection   : .\Invoke-BitLockerEscrowCheck.ps1 -DetectOnly
      remediation : .\Invoke-BitLockerEscrowCheck.ps1

.PARAMETER DetectOnly
    Only report. Changes nothing. Exit code 1 if something is not in order.

.PARAMETER AddRecoveryPasswordIfMissing
    Adds a RecoveryPassword protector to encrypted drives that have none,
    and then backs it up to Entra ID.

.PARAMETER IncludeRemovable
    Also checks removable drives (USB sticks and similar).

.EXAMPLE
    .\Invoke-BitLockerEscrowCheck.ps1 -DetectOnly

.EXAMPLE
    .\Invoke-BitLockerEscrowCheck.ps1 -AddRecoveryPasswordIfMissing -WhatIf

.NOTES
    Exit codes: 0 = everything in order, 1 = problems remain (unencrypted drive,
    suspended protection or escrow missing/failed), 2 = fatal error or not Entra joined.
    Log: %ProgramData%\BitLockerCheck\
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$DetectOnly,
    [switch]$AddRecoveryPasswordIfMissing,
    [switch]$IncludeRemovable,
    [string]$LogDirectory = (Join-Path $env:ProgramData 'BitLockerCheck')
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

function Test-EntraJoined {
    $status = & "$env:windir\System32\dsregcmd.exe" /status 2>$null
    return [bool]($status -match '^\s*AzureAdJoined\s*:\s*YES')
}

function Get-EscrowedProtectorIds {
    <#
        Returns the IDs (without braces, upper case) of protectors for which event 845
        ("backed up successfully to your Azure AD") exists.
    #>
    $ids = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $events = Get-WinEvent -FilterHashtable @{
        LogName = 'Microsoft-Windows-BitLocker/BitLocker Management'
        Id      = 845
    } -ErrorAction SilentlyContinue

    foreach ($evt in @($events)) {
        if (-not $evt) { continue }
        # Search both the message and the raw event properties for GUIDs
        $text = ($evt.Message, ($evt.Properties | ForEach-Object { $_.Value })) -join ' '
        foreach ($m in [regex]::Matches($text, '[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}')) {
            [void]$ids.Add($m.Value)
        }
    }
    return ,$ids
}

function Get-TargetVolumes {
    $fixedLetters = Get-CimInstance -ClassName Win32_LogicalDisk |
        Where-Object { $_.DriveType -eq 3 -or ($IncludeRemovable -and $_.DriveType -eq 2) } |
        ForEach-Object { $_.DeviceID }

    Get-BitLockerVolume | Where-Object {
        $_.VolumeType -eq 'OperatingSystem' -or $fixedLetters -contains $_.MountPoint
    }
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('BitLocker_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'BitLocker_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    $mode = if ($DetectOnly) { 'DetectOnly' } elseif ($WhatIfPreference) { 'WhatIf' } else { 'Remediate' }
    Write-Log "=== BitLocker check started on $env:COMPUTERNAME ($mode) ==="

    if (-not (Test-EntraJoined)) {
        Write-Log 'Device is not Entra ID joined; backup to Entra ID is not possible.' 'ERROR'
        exit 2
    }

    $escrowed = Get-EscrowedProtectorIds
    Write-Log "Protectors with an escrow event (845) in the log: $($escrowed.Count)"

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($vol in @(Get-TargetVolumes)) {
        $mp = $vol.MountPoint
        $result = [ordered]@{
            Drive      = $mp
            Type       = "$($vol.VolumeType)"
            Status     = "$($vol.VolumeStatus)"
            Protection = "$($vol.ProtectionStatus)"
            Escrow     = 'n/a'
            Compliant  = $true
        }
        Write-Log ("Drive {0}: {1}, protection {2}, {3}% encrypted" -f `
            $mp, $vol.VolumeStatus, $vol.ProtectionStatus, $vol.EncryptionPercentage)

        # 1. Not encrypted -> report only
        if ($vol.VolumeStatus -eq 'FullyDecrypted') {
            Write-Log "  $mp is NOT encrypted." 'WARN'
            $result.Compliant = $false
            $results.Add([pscustomobject]$result)
            continue
        }

        if ($vol.ProtectionStatus -ne 'On' -and $vol.VolumeStatus -eq 'FullyEncrypted') {
            Write-Log "  $mp is encrypted but protection is suspended/off." 'WARN'
            $result.Compliant = $false
        }

        # 2. Make sure a RecoveryPassword protector exists
        $recovery = @($vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })

        if ($recovery.Count -eq 0) {
            Write-Log "  $mp has no RecoveryPassword protector." 'WARN'
            if ($AddRecoveryPasswordIfMissing -and -not $DetectOnly) {
                if ($PSCmdlet.ShouldProcess($mp, 'Add RecoveryPassword protector')) {
                    try {
                        Add-BitLockerKeyProtector -MountPoint $mp -RecoveryPasswordProtector `
                            -WarningAction SilentlyContinue | Out-Null
                        $recovery = @((Get-BitLockerVolume -MountPoint $mp).KeyProtector |
                            Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })
                        Write-Log "  RecoveryPassword protector added to $mp."
                    }
                    catch { Write-Log "  Adding protector failed: $($_.Exception.Message)" 'ERROR' }
                }
            }
            if ($recovery.Count -eq 0) {
                $result.Escrow    = 'No recovery key'
                $result.Compliant = $false
                $results.Add([pscustomobject]$result)
                continue
            }
        }

        # 3. Back up every protector that has no escrow event
        $missing = @($recovery | Where-Object { -not $escrowed.Contains($_.KeyProtectorId.Trim('{}')) })
        if ($missing.Count -eq 0) {
            Write-Log "  All recovery keys on $mp are escrowed to Entra ID."
            $result.Escrow = 'OK'
            $results.Add([pscustomobject]$result)
            continue
        }

        $failed = 0
        foreach ($kp in $missing) {
            $id = $kp.KeyProtectorId
            Write-Log "  Protector $id on ${mp}: no escrow found." 'WARN'
            if ($DetectOnly) { $failed++; continue }

            if ($PSCmdlet.ShouldProcess("$mp protector $id", 'BackupToAAD-BitLockerKeyProtector')) {
                try {
                    BackupToAAD-BitLockerKeyProtector -MountPoint $mp -KeyProtectorId $id | Out-Null
                    Write-Log "  Protector $id backed up to Entra ID."
                }
                catch {
                    Write-Log "  Backup of $id failed: $($_.Exception.Message)" 'ERROR'
                    $failed++
                }
            }
            else { $failed++ }   # WhatIf: not actually backed up
        }

        if ($failed -gt 0) {
            $result.Escrow    = "Missing ($failed)"
            $result.Compliant = $false
        }
        else { $result.Escrow = 'Backed up now' }
        $results.Add([pscustomobject]$result)
    }

    if ($results.Count -eq 0) { Write-Log 'No drives found to check.' 'WARN' }

    $table = $results | Format-Table -AutoSize | Out-String
    Write-Log ("Summary:`r`n" + $table.TrimEnd())

    $nonCompliant = @($results | Where-Object { -not $_.Compliant })
    if ($nonCompliant.Count -gt 0) {
        # Single line on stdout for Intune "pre/post-remediation detection output"
        Write-Output ("NON-COMPLIANT: " + (($nonCompliant | ForEach-Object { "$($_.Drive) $($_.Status)/$($_.Escrow)" }) -join '; '))
        exit 1
    }
    Write-Output 'COMPLIANT: all drives encrypted and recovery keys escrowed to Entra ID'
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
