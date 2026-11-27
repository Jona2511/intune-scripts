# Invoke-WindowsUpdateRepair

Detects a stuck Windows Update and resets the update components.

## Description

```text
Detection:
  - Date of the last successfully installed update, taken from the Windows Update
    history (Microsoft.Update.Session). By default only updates with a KB number count;
    Defender definitions, MSRT and Store app updates are ignored because they install
    often and would hide a stuck cumulative update.
    Fallback when the history is empty: Get-HotFix.
  - Whether the wuauserv service is disabled.
  - Whether a reboot is pending. In that case a reset is pointless: the script reports
    it and skips the reset (unless -IgnorePendingReboot).

Reset (when the last update is older than -MaxDaysSinceUpdate):
  1. Stops wuauserv, bits, cryptsvc, msiserver and usosvc.
  2. Deletes the BITS queue (qmgr*.dat).
  3. Renames SoftwareDistribution and catroot2 to *.bak_<timestamp>
     (older .bak folders from this script are removed).
  4. Sets wuauserv back to Manual if it was disabled.
  5. Starts the services again and triggers a new scan.
  Optional with -RunDism: DISM /RestoreHealth and SFC /scannow (takes a long time).

Resetting clears the visible update history in Settings. To avoid resetting every
run, the script waits at least -MinDaysBetweenResets after a previous reset.
```

## Parameters

- `-MaxDaysSinceUpdate`: The reset runs if the last successful update is older than this (default 35).
- `-MinDaysBetweenResets`: Minimum number of days between two resets (default 7).
- `-CountAllUpdates`: Count every successful update, including Defender definitions, Store apps and drivers. By default only updates with a KB number count (Windows, .NET, Office), excluding Defender/MSRT updates.
- `-IgnorePendingReboot`: Resets even when a reboot is pending.
- `-RunDism`: Also runs DISM /RestoreHealth and SFC /scannow after the reset.
- `-Force`: Resets regardless of the update age and the time since the last reset.
- `-DetectOnly`: Only report. Exit code 1 if Windows Update looks stuck.

## Examples

```powershell
.\Invoke-WindowsUpdateRepair.ps1 -DetectOnly
```

```powershell
.\Invoke-WindowsUpdateRepair.ps1 -MaxDaysSinceUpdate 45 -RunDism
```

## Notes

```text
Exit codes: 0 = in order or repaired, 1 = stuck / reboot needed / reset failed,
2 = fatal error.
Log: %ProgramData%\WindowsUpdateRepair\
```
