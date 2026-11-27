# Invoke-StartTaskbarCleanup

Cleans up the taskbar and Start menu on new Windows 11 devices by removing the
standard Microsoft pins.

## Description

```text
What it configures (machine level, applies to NEW user profiles):

  - Taskbar: writes C:\Windows\OEM\TaskbarLayoutModification.xml with
    PinListPlacement="Replace" (only File Explorer plus -ExtraTaskbarPins) and registers
    it via HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\LayoutXMLPath.
    Windows applies it once, at a user's first sign-in. Users can change pins afterwards.

  - Start menu: copies a start2.bin (made on a reference machine with
    -ExportStartLayout) to the Default profile:
    C:\Users\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState
    New users get that pin layout at first sign-in.

  - Optional:
      -HideWidgets      policy  HKLM\SOFTWARE\Policies\Microsoft\Dsh\AllowNewsAndInterests = 0
      -SearchMode       policy  HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\SearchOnTaskbarMode
      -HideTaskView     Default user hive: Explorer\Advanced\ShowTaskViewButton = 0

  - Optional -ResetExistingUsers (best effort): for existing profiles, removes the
    taskbar pin state (Taskband key and User Pinned\TaskBar shortcuts) so the layout file
    is used again, and copies start2.bin into their profile. Explorer and the Start menu
    are restarted for users who are signed in.

What is NOT supported / limitations (honest version):
  - Windows 11 has no API to unpin individual Start or taskbar items. This script
    replaces the whole layout instead.
  - start2.bin is an undocumented binary format. It is copied as-is and can break after
    a feature update; re-export it from a reference machine on the same build.
  - Several HKCU taskbar values (e.g. TaskbarDa for Widgets) are protected by the UCPD
    driver on recent builds and cannot be set by scripts. That is why Widgets and Search
    use policies here instead of HKCU values.
  - The taskbar layout is applied once per user. -ResetExistingUsers relies on the fact
    that Explorer rebuilds pins from the layout file when the pin state is missing; this
    works on current builds but is not documented behaviour.
  - Store-app pins (Store, Outlook, Copilot) live only in the Taskband state, so they
    disappear only through the layout file or a reset, not by deleting shortcuts.
```

## Parameters

- `-StartLayoutBin`: Path to a start2.bin exported with -ExportStartLayout. Without it, the Start menu is skipped.
- `-ExtraTaskbarPins`: Extra taskbar pins next to File Explorer. Per entry: - contains '!' -> Store app AUMID, e.g. 'MSTeams_8wekyb3d8bbwe!MSTeams' - ends with '.lnk' -> shortcut path, e.g. '%APPDATA%\Microsoft\Windows\Start Menu\Programs\Outlook.lnk' - otherwise -> desktop application ID, e.g. 'MSEdge'
- `-ExportStartLayout`: Exports the current user's start2.bin to -ExportPath and stops. Run this as the user on a reference machine after arranging the Start menu manually.
- `-SearchMode`: Hide, Icon, IconAndLabel or Box. Not set = leave as is.

## Examples

```powershell
.\Invoke-StartTaskbarCleanup.ps1 -ExportStartLayout -ExportPath C:\Temp\start2.bin
```

```powershell
.\Invoke-StartTaskbarCleanup.ps1 -StartLayoutBin .\start2.bin -HideWidgets -SearchMode Icon -HideTaskView
```

```powershell
.\Invoke-StartTaskbarCleanup.ps1 -StartLayoutBin .\start2.bin -DetectOnly
```

## Notes

```text
Exit codes: 0 = configured, 1 = not (fully) configured, 2 = fatal error.
Log: %ProgramData%\StartTaskbarCleanup\
```
