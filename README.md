# intune-scripts

PowerShell scripts I use for Microsoft Intune and Microsoft 365: packaging apps, remediations, platform scripts, custom compliance, Autopilot, user management and automation through Microsoft Graph.

## Structure

| Folder | Contents |
|---|---|
| [Apps](Apps) | Win32 apps, one folder per app with install, uninstall and detection scripts |
| [Remediations](Remediations) | Detection and remediation pairs |
| [Platform-Scripts](Platform-Scripts) | Scripts that run once on a device |
| [Compliance](Compliance) | Custom compliance: discovery scripts with their JSON rules |
| [Autopilot](Autopilot) | Device registration, group tags and other Autopilot tasks |
| [Graph](Graph) | Reporting and bulk changes through Microsoft Graph |
| [Users](Users) | Onboarding, offboarding and cleaning up user accounts |
| [Templates](Templates) | Starting points for new scripts |

## Before you use a script

Read it first and try it on a test device or a small test group. Unless the header says otherwise, the scripts expect to run as SYSTEM in 64-bit PowerShell, which is how Intune runs them when you set "Run script in 64 bit PowerShell Host" to Yes.

Scripts that run on devices log to `%ProgramData%\<ScriptName>\`. Most of the scripts that change something support `-WhatIf`, so you can see what they would do first.

## Conventions

Every script has its own folder with a README that lists its parameters, examples and, for Graph scripts, the permissions it needs. The scripts start with comment-based help, so `Get-Help .\Script.ps1 -Full` works too.

Remediation scripts do detection and repair in one file: `-DetectOnly` only checks and exits with 1 when something needs fixing, without it they repair. See [Remediations](Remediations) for the two exceptions.

Nothing here contains tenant IDs, customer names or secrets. Values that differ per tenant are parameters or variables at the top of the script.

A GitHub Action runs [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) on every push and pull request.

## License

[MIT](LICENSE)
