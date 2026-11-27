# Test-HuntressAgent

Custom compliance check: is Huntress installed, and are its services running?

## What it returns

| Setting | Type | Compliant when |
|---|---|---|
| `HuntressInstalled` | Boolean | true: at least one service with a display name like `Huntress*` exists |
| `HuntressRunning` | Boolean | true: every Huntress service set to start automatically is running |

Services set to start manually are left out, because those don't have to run all the time.

## Setup in Intune

1. Devices > Compliance > Scripts: add `Test-HuntressAgent.ps1` as a Windows discovery script (run as system, 64-bit).
2. In a Windows compliance policy, under Custom Compliance, pick the script and upload `Test-HuntressAgent.json`.

If your Huntress services use a different display name, change `$ServicePattern` at the top of the script.
