# Invoke-OneDrivePerMachineCheck

Detects per-user OneDrive installs and moves the device to the per-machine client.

## Description

```text
Detection:
  - Per-machine install: %ProgramFiles%\Microsoft OneDrive\OneDrive.exe
    (or %ProgramFiles(x86)% on older installs).
  - Per-user installs: <profile>\AppData\Local\Microsoft\OneDrive\OneDrive.exe for
    every user profile.
  - Optional: per-machine version older than -MinimumVersion.

Remediation (unless -DetectOnly):
  1. Per-machine client missing or too old -> runs OneDriveSetup.exe /allusers /silent.
     The installer comes from -InstallerPath (local/UNC) or is downloaded from -InstallerUrl.
  2. For every profile with a per-user install, the user's Run key (HKCU\...\Run\OneDrive)
     is pointed to the per-machine OneDrive.exe, so the user starts the per-machine
     client at the next sign-in. Hives of users who are not signed in are loaded
     temporarily.
  3. With -RemovePerUserBinaries: removes the per-user program files (OneDrive.exe,
     OneDriveStandaloneUpdater.exe and the version folders), the per-user uninstall
     entry and the per-user update task. Sync settings and data (the 'settings' folder
     and the OneDrive folder with files) are never touched. Skipped for users who have
     OneDrive running at that moment.
```

## Parameters

- `-InstallerPath`: Local or UNC path to OneDriveSetup.exe. Takes precedence over -InstallerUrl.
- `-InstallerUrl`: Download URL for OneDriveSetup.exe.
- `-MinimumVersion`: Optional: minimum per-machine version, e.g. '25.100.0.0'.
- `-RemovePerUserBinaries`: Also removes the per-user program files (see description).
- `-DetectOnly`: Only report. Exit code 1 if per-user installs exist or the per-machine client is missing.

## Examples

```powershell
.\Invoke-OneDrivePerMachineCheck.ps1 -DetectOnly
```

```powershell
.\Invoke-OneDrivePerMachineCheck.ps1 -InstallerPath '\\fs01\sw$\OneDrive\OneDriveSetup.exe' -RemovePerUserBinaries
```

## Notes

```text
Exit codes: 0 = per-machine only (afterwards), 1 = per-user installs remain / install failed,
2 = fatal error.
Log: %ProgramData%\OneDriveCheck\
```
