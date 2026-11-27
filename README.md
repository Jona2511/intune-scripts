# intune-scripts

PowerShell scripts I use for Microsoft Intune: packaging apps, remediations, platform scripts, custom compliance, Autopilot and automation through Microsoft Graph.

## Structure

| Folder | Contents |
|---|---|
| [Apps](Apps) | Win32 apps, one folder per app with install, uninstall and detection scripts |
| [Remediations](Remediations) | Detection and remediation pairs |
| [Platform-Scripts](Platform-Scripts) | Scripts that run once on a device |
| [Compliance](Compliance) | Custom compliance: discovery scripts with their JSON rules |
| [Autopilot](Autopilot) | Device registration, group tags and other Autopilot tasks |
| [Graph](Graph) | Reporting and bulk changes through Microsoft Graph |
| [Templates](Templates) | Starting points for new scripts |

## Before you use a script

Read it first and try it on a test device or a small test group. Unless the header says otherwise, the scripts expect to run as SYSTEM in 64-bit PowerShell, which is how Intune runs them when you set "Run script in 64 bit PowerShell Host" to Yes.

Scripts log to `C:\ProgramData\IntuneScripts\Logs`, one log file per script.

## Conventions

Every script starts with comment-based help (`.SYNOPSIS`, `.DESCRIPTION` and `.NOTES` with the version and the context it runs in), so `Get-Help .\Script.ps1` works.

Remediations come in pairs named `Detect-<Name>.ps1` and `Remediate-<Name>.ps1`. The detection script exits with 1 when something needs fixing and 0 when the device is fine.

Nothing here contains tenant IDs, customer names or secrets. Values that differ per tenant are parameters or variables at the top of the script.

A GitHub Action runs [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) on every push and pull request.

## License

[MIT](LICENSE)
