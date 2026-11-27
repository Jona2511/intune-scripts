# Invoke-LowDiskCleanup

Cleans up C: when free space drops below a threshold.

## Description

```text
Checks free space on C:. If it is below -MinFreeGB, the script runs these steps,
from least to most invasive, and stops once enough space is free again:

  1. Temp folders     : C:\Windows\Temp and every user's AppData\Local\Temp
                        (only files older than -TempFileAgeDays).
  2. Windows Update   : SoftwareDistribution\Download and the Delivery Optimization cache.
                        wuauserv and BITS are stopped and restarted afterwards.
  3. Old user profiles: local profiles not used for -ProfileAgeDays days, removed through
                        Win32_UserProfile (folder and registry entry). Profiles that are
                        loaded, special, the current user or in -ExcludeProfiles are skipped.

Supports -WhatIf: shows what would be deleted without changing anything.
```

## Parameters

- `-MinFreeGB`: Threshold in GB. The cleanup runs only below this value (unless -Force).
- `-ProfileAgeDays`: Profiles not used for this many days are considered old.
- `-TempFileAgeDays`: Only temp files older than this many days are deleted (0 = everything).
- `-ExcludeProfiles`: Profile folder names (wildcards allowed) that are never removed.
- `-Force`: Runs all steps regardless of free space.

## Examples

```powershell
.\Invoke-LowDiskCleanup.ps1 -MinFreeGB 15 -WhatIf
Shows what would be cleaned up if C: has less than 15 GB free.
```

```powershell
.\Invoke-LowDiskCleanup.ps1 -MinFreeGB 20 -ProfileAgeDays 60 -ExcludeProfiles 'admin*','svc_*'
```

## Notes

```text
Exit codes: 0 = enough free space (afterwards), 1 = still below threshold, 2 = fatal error.
Log: %ProgramData%\DiskCleanup\
```
