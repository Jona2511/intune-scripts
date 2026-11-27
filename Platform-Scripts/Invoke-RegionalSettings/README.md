# Invoke-RegionalSettings

Sets the time zone to W. Europe Standard Time, Dutch regional formats and home location,
and keeps the keyboard on US International only.

## Description

```text
Target state:
  - Time zone            : W. Europe Standard Time (machine)
  - Regional format      : nl-NL (dates, numbers, currency) per user
  - Home location        : Netherlands (GeoId 176) per user
  - Keyboard             : ONLY US International (00020409) for every language in the
                           user's language list (e.g. 0409:00020409, 0413:00020409)
  - Display language     : not changed (the language list order stays as it is)
  - System locale        : nl-NL only with -SetSystemLocale (needs a reboot)

The script works in two modes:

  - Machine mode (SYSTEM or elevated admin; default):
      1. Sets the time zone (and system locale with -SetSystemLocale).
      2. Applies the user settings to its own context and, if available (Windows 11),
         copies them with Copy-UserInternationalSettingsToSystem to the welcome screen
         and to new user profiles.
      3. Runs the per-user part for signed-in users through a temporary scheduled task
         for BUILTIN\Users (the script copies itself to %ProgramData%\RegionalSettings).
         Users who are not signed in get the settings at the next run, or via step 2 if
         their profile is new.

  - User mode (-UserMode, used by the scheduled task): applies format, home location and
    keyboard for the current user. Runs without elevation.

Detection checks the time zone, the system locale (if requested), the welcome-screen
defaults (HKU\.DEFAULT) and the settings of every signed-in user (via HKU).
```

## Parameters

- `-TimeZone`: Time zone ID (default 'W. Europe Standard Time').
- `-Culture`: Regional format (default 'nl-NL').
- `-GeoId`: Home location (default 176 = Netherlands).
- `-KeyboardLayout`: Keyboard layout ID kept for every language (default '00020409' = US International).
- `-SetSystemLocale`: Also sets the system locale (language for non-Unicode programs) to -Culture. Needs a reboot.
- `-DetectOnly`: Only report. Exit code 1 if something differs.
- `-UserMode`: Internal: applies the per-user settings in the current user session.

## Examples

```powershell
.\Invoke-RegionalSettings.ps1 -DetectOnly
```

```powershell
.\Invoke-RegionalSettings.ps1 -SetSystemLocale
```

## Notes

```text
Exit codes: 0 = in order (afterwards), 1 = differences remain, 2 = fatal error.
Log (machine): %ProgramData%\RegionalSettings\
Log (user)   : %LOCALAPPDATA%\RegionalSettings\
Run in machine mode as an elevated admin, Copy-UserInternationalSettingsToSystem copies
that admin's own settings (including the language list) to the welcome screen and new users.
```
