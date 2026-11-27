# Test-ForbiddenApps

Custom compliance check: counts installed apps that are not allowed, such as remote control or torrent software.

## What it returns

| Setting | Type | Compliant when |
|---|---|---|
| `ForbiddenAppCount` | Int64 | 0 |

The script checks the uninstall entries for the whole device (64-bit and 32-bit) and for every user who is signed in, so per-user installs are found too. Which apps it found goes to `%ProgramData%\Test-ForbiddenApps\Test-ForbiddenApps.log`, because a compliance script may only output its JSON.

## The list

Edit `$ForbiddenApps` at the top of the script. Wildcards work:

```powershell
$ForbiddenApps = @(
    'TeamViewer*'
    'AnyDesk*'
    '*uTorrent*'
    '*BitTorrent*'
)
```

If your own remote support tool matches one of these (a company TeamViewer, for example), take it off the list.

## Setup in Intune

1. Devices > Compliance > Scripts: add `Test-ForbiddenApps.ps1` as a Windows discovery script (run as system, 64-bit).
2. In a Windows compliance policy, under Custom Compliance, pick the script and upload `Test-ForbiddenApps.json`.
