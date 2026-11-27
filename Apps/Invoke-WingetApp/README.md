# Invoke-WingetApp

Universal winget template: install, upgrade, uninstall or detect any app as SYSTEM.

## Description

```text
Built for Intune Win32 apps (and RMM) that run as SYSTEM, where 'winget' is not on PATH.

  - Resolves winget.exe from %ProgramFiles%\WindowsApps\Microsoft.DesktopAppInstaller_*
    (newest version, matching CPU architecture). Falls back to 'winget' on PATH when
    running as a user.
  - Detection: 'winget list --id <AppId> --exact'. The line containing the exact Id is
    parsed by tokens (no column headers, so locale independent). If winget truncated
    the Id (ellipsis), 'winget export --include-versions' is used instead. Optional extra
    fallback: uninstall registry keys via -DisplayNamePattern.
  - Install: 'winget install' (upgrades an existing older install). If the package has
    no installer for the requested scope, it retries without --scope.
  - Intune detection mode (-DetectOnly): writes ONE line to stdout and exits 0 only when
    the app is installed with a version >= the required version; otherwise writes
    nothing to stdout (log file only).

Intune Win32 app command lines (package this script with IntuneWinAppUtil):
  Install  : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-WingetApp.ps1 -AppId 7zip.7zip
  Uninstall: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-WingetApp.ps1 -AppId 7zip.7zip -Uninstall
  Detection: custom script = a copy of this script with the default values of
             -AppId / -MinimumVersion filled in and -DetectOnly set to $true
             (Intune detection scripts cannot take parameters).
  Return codes: 0 = success, 3010 = soft reboot, 1 = failed.
```

## Parameters

- `-AppId`: winget package identifier, e.g. '7zip.7zip' or 'Microsoft.PowerShell'.
- `-Version`: Install exactly this version (pin). Detection then requires at least this version.
- `-MinimumVersion`: Detection/compliance requires at least this version. Without -Version or -MinimumVersion any installed version counts as detected.
- `-Scope`: 'machine' (default), 'user' or 'none' (no --scope argument).
- `-Override`: Replaces all installer arguments (winget --override).
- `-Custom`: Appends arguments to the default silent arguments (winget --custom).
- `-Source`: winget source (default 'winget').
- `-DisplayNamePattern`: Optional wildcard for the uninstall registry DisplayName, used as detection fallback when winget cannot be run or cannot find the package.
- `-Uninstall`: Uninstalls the package.
- `-DetectOnly`: Intune detection mode (see description).

## Examples

```powershell
.\Invoke-WingetApp.ps1 -AppId 'Notepad++.Notepad++' -MinimumVersion 8.6
```

```powershell
.\Invoke-WingetApp.ps1 -AppId 'Zoom.Zoom' -Custom 'ZoomAutoUpdate=1'
```

## Notes

```text
Exit codes (install/uninstall): 0 = success, 3010 = reboot required, 1 = failed.
Exit codes (-DetectOnly)      : 0 + stdout = detected, 1 + no stdout = not detected.
Log: %ProgramData%\WingetApp\<AppId>_*.log  (winget's own log: <AppId>_winget_*.log)
Under SYSTEM, winget needs the Visual C++ 2015-2022 x64 runtime; without it winget.exe
fails with 0xC0000135.
```
