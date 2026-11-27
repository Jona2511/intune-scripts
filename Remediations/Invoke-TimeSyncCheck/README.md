# Invoke-TimeSyncCheck

Checks the Windows Time service and clock drift, and repairs them if needed.

## Description

```text
Checks:
  - Whether the Windows Time service (W32Time) is disabled or not running.
  - Which time source the client uses. 'Local CMOS Clock' or 'Free-running System
    Clock' means it is not syncing with anything.
  - The clock offset compared with a reference server, measured with
    'w32tm /stripchart'. The median of several samples is used.

Repair (unless -DetectOnly):
  - Disabled service -> startup type set to Automatic.
  - Service stopped or drift above -MaxDriftSeconds -> service restarted,
    then 'w32tm /resync /force'. If that fails: 'w32tm /config /update' and
    'w32tm /resync /rediscover'.
  - With -FixSource on a non-domain machine without a valid source: sets -NtpServer
    as the NTP source. Domain machines always keep the domain hierarchy.
  - Measures the drift again afterwards.
```

## Parameters

- `-MaxDriftSeconds`: Maximum allowed offset in seconds (default 5). Kerberos fails above 300.
- `-ReferenceServer`: Server to measure the drift against. Default: the current time source, or -NtpServer if there is no valid source.
- `-NtpServer`: NTP server used as a fallback reference and for -FixSource.
- `-FixSource`: Configures -NtpServer as the time source on non-domain machines without a valid source.
- `-DetectOnly`: Only report. Exit code 1 if something is wrong.

## Examples

```powershell
.\Invoke-TimeSyncCheck.ps1 -DetectOnly
```

```powershell
.\Invoke-TimeSyncCheck.ps1 -MaxDriftSeconds 2 -FixSource -NtpServer 'nl.pool.ntp.org'
```

## Notes

```text
Exit codes: 0 = in order (afterwards), 1 = problem remains, 2 = fatal error.
Log: %ProgramData%\TimeSyncCheck\
'w32tm /stripchart' needs outbound UDP 123 to the reference server.
```
