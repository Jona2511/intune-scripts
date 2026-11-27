# Invoke-UptimeWarning

Detects devices that have not rebooted for more than X days and shows the
logged-on user a toast notification.

## Description

```text
Uptime is based on Win32_OperatingSystem.LastBootUpTime. With Fast Startup enabled
a "Shut down" does not reset the uptime; only "Restart" does. The toast says so.

The script works in two contexts:

  - As SYSTEM (Intune remediation, scheduled task, RMM):
    A toast cannot be shown from the SYSTEM session. The script copies itself to
    %ProgramData%\UptimeWarning, writes the toast texts to config.json and starts a
    temporary scheduled task that runs in the session of every logged-on user
    (BUILTIN\Users, interactive). That task runs this script with -ShowToast.
    The task is removed afterwards.

  - As the logged-on user:
    Shows the toast directly.

The toast has two buttons:
  - "Nu herstarten": restarts after 60 seconds (cancel with 'shutdown /a').
    Uses a protocol handler (uptime-restart:) in HKCU of the user.
  - "Later": closes the notification.
```

## Parameters

- `-MaxUptimeDays`: Warn when the device has been up longer than this (default 14).
- `-Title`: Toast title.
- `-Message`: Toast text. {0} is replaced by the number of days.
- `-DetectOnly`: Only report, show no toast. Exit code 1 if the uptime is too long.
- `-ShowToast`: Internal: shows the toast in the current user session (used by the scheduled task).

## Examples

```powershell
.\Invoke-UptimeWarning.ps1 -DetectOnly
```

```powershell
.\Invoke-UptimeWarning.ps1 -MaxUptimeDays 7
```

## Notes

```text
Exit codes: 0 = uptime OK, 1 = uptime too long (toast shown or not), 2 = fatal error.
Log (SYSTEM): %ProgramData%\UptimeWarning\
Log (user)  : %LOCALAPPDATA%\UptimeWarning\
```
