# Invoke-BitLockerEscrowCheck

Checks BitLocker status and makes sure recovery keys are escrowed to Entra ID.

## Description

```text
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
```

## Parameters

- `-DetectOnly`: Only report. Changes nothing. Exit code 1 if something is not in order.
- `-AddRecoveryPasswordIfMissing`: Adds a RecoveryPassword protector to encrypted drives that have none, and then backs it up to Entra ID.
- `-IncludeRemovable`: Also checks removable drives (USB sticks and similar).

## Examples

```powershell
.\Invoke-BitLockerEscrowCheck.ps1 -DetectOnly
```

```powershell
.\Invoke-BitLockerEscrowCheck.ps1 -AddRecoveryPasswordIfMissing -WhatIf
```

## Notes

```text
Exit codes: 0 = everything in order, 1 = problems remain (unencrypted drive,
suspended protection or escrow missing/failed), 2 = fatal error or not Entra joined.
Log: %ProgramData%\BitLockerCheck\
```
