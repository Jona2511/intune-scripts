# Invoke-BloatwareRemoval

Removes provisioned Windows apps (bloatware) that are not on an allowlist.

## Description

```text
- Reads all provisioned Appx packages (Get-AppxProvisionedPackage -Online). These are
  installed for every new user profile.
- Every package whose DisplayName does not match the allowlist is removed with
  Remove-AppxProvisionedPackage, so new users no longer get it.
- Framework packages (VCLibs, UI.Xaml, .NET Native, WindowsAppRuntime) and packages
  Windows marks as NonRemovable are always skipped.
- With -IncludeInstalled the same apps are also removed for existing users
  (Get-AppxPackage -AllUsers | Remove-AppxPackage -AllUsers). Only Microsoft-published
  packages (and the provisioned apps found above) are considered there, so apps users
  installed themselves from other publishers are left alone.

Default allowlist (wildcards on the package name): Store, Store purchase, App Installer
(winget), Terminal, Calculator, Photos, Notepad, Paint, Snipping Tool, Camera, Windows
Security, codec extensions (HEIF, HEVC, VP9, WebP, AV1, Raw, MPEG2, Web Media), Company
Portal, new Teams, Quick Assist, Edge, winget source and framework/runtime packages.
Extend it with -AdditionalAllowList or -AllowListPath, or replace it with -AllowList.
```

## Parameters

- `-AllowList`: Replaces the default allowlist.
- `-AdditionalAllowList`: Added to the (default) allowlist, e.g. 'Microsoft.ZuneMusic','Microsoft.OutlookForWindows'.
- `-AllowListPath`: Text file with extra allowed names, one per line ('#' = comment).
- `-IncludeInstalled`: Also removes the apps for existing user profiles.
- `-DetectOnly`: Only report. Exit code 1 if apps outside the allowlist are present.

## Examples

```powershell
.\Invoke-BloatwareRemoval.ps1 -DetectOnly
```

```powershell
.\Invoke-BloatwareRemoval.ps1 -AdditionalAllowList 'Microsoft.ZuneMusic' -IncludeInstalled -WhatIf
```

## Notes

```text
Exit codes: 0 = nothing (left) to remove, 1 = apps remain / removal failed, 2 = fatal error.
Log: %ProgramData%\BloatwareRemoval\
Removed apps can be reinstalled from the Microsoft Store or winget.
```
