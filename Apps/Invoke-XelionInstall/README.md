# Invoke-XelionInstall

Silent install, uninstall and Intune detection for the Xelion softphone (Windows app).

## Description

```text
Xelion distributes its Windows app through the Microsoft Store and apps.xelion.com.
Verified on a reference device: the current Windows app is an MSIX/Store package
named '71F23AFD.Xelion' (publisher id c91qzwq3fg5hw, Store Id 9NBLGGH5Q9BJ).
The separate "Xelion Admin" tool is a per-user ClickOnce app and is excluded from
detection by default (-ExcludeDisplayNamePattern).

The installer format can differ per release, so this script supports all three
common formats and picks the method from the file extension:

  - MSIX / MSIXBUNDLE / APPX / APPXBUNDLE:
      Provisioned for all users with Add-AppxProvisionedPackage -SkipLicense.
      Existing users get the app at their next sign-in; new users at first sign-in.
  - MSI:
      msiexec /i <file> /qn /norestart /l*v <log>  (+ -MsiProperties)
  - EXE:
      <file> -SilentArgs (no universal silent switch exists for EXE installers:
      check the vendor's switches and set -SilentArgs).

Without -InstallerPath the script uses the first Xelion* installer it finds next to
itself (the Intune .intunewin content folder).

Detection (installed + version >= -MinimumVersion), all of these are checked:
  - provisioned and installed AppX/MSIX packages whose name matches -AppxNamePattern
  - uninstall registry entries (HKLM 64/32-bit and the loaded HKU user hives)
    whose DisplayName matches -DisplayNamePattern

Store alternative (no installer file needed): Intune "Microsoft Store app (new)" with
Id 9NBLGGH5Q9BJ, or Invoke-WingetApp.ps1 -AppId 9NBLGGH5Q9BJ -Source msstore.
This script's -DetectOnly detects that Store install as well.

Intune Win32 app command lines:
  Install  : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-XelionInstall.ps1
  Uninstall: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-XelionInstall.ps1 -Uninstall
  Detection: custom script = copy of this script with -DetectOnly / -MinimumVersion
             defaults filled in (detection scripts cannot take parameters).
```

## Parameters

- `-InstallerPath`: Path to the Xelion installer (.msix/.msixbundle/.appx/.appxbundle/.msi/.exe).
- `-SilentArgs`: Arguments for an .exe installer (default '/S' - verify for your installer version).
- `-MsiProperties`: Extra MSI properties, e.g. 'ALLUSERS=1'.
- `-MinimumVersion`: Minimum version for detection/compliance. Empty = any version.
- `-CloseProcesses`: Process names (wildcards) to close before install/uninstall. Empty array = do not close.
- `-AppxNamePattern`: Wildcard for the AppX/MSIX package name.
- `-DisplayNamePattern`: Wildcard for the uninstall registry DisplayName.
- `-ExcludeDisplayNamePattern`: DisplayName wildcards to ignore (default: 'Xelion Admin*').
- `-Uninstall`: Removes Xelion (AppX/MSIX for all users + deprovision, MSI, or EXE uninstaller).
- `-UninstallArgs`: Arguments for an EXE uninstaller without QuietUninstallString (default '/S').
- `-DetectOnly`: Intune detection mode: stdout + exit 0 when detected, nothing on stdout otherwise.

## Examples

```powershell
.\Invoke-XelionInstall.ps1 -InstallerPath .\XelionSetup.msix -MinimumVersion 8.4
```

```powershell
.\Invoke-XelionInstall.ps1 -DetectOnly -MinimumVersion 8.4
```

## Notes

```text
Exit codes: 0 = success, 3010 = reboot required, 1 = failed.
-DetectOnly: 0 + stdout = detected, 1 + no stdout = not detected.
Run as SYSTEM/administrator. Log: %ProgramData%\XelionInstall\
```
