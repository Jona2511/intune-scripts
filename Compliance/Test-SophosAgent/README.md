# Test-SophosAgent

Custom compliance check: is Sophos installed, and are its services running?

## What it returns

| Setting | Type | Compliant when |
|---|---|---|
| `SophosInstalled` | Boolean | true: at least one service with a display name like `Sophos*` exists |
| `SophosRunning` | Boolean | true: every Sophos service set to start automatically is running |

Services set to start manually are left out, because those don't have to run all the time.

## Setup in Intune

1. Devices > Compliance > Scripts: add `Test-SophosAgent.ps1` as a Windows discovery script (run as system, 64-bit).
2. In a Windows compliance policy, under Custom Compliance, pick the script and upload `Test-SophosAgent.json`.

If your Sophos services use a different display name, change `$ServicePattern` at the top of the script.
