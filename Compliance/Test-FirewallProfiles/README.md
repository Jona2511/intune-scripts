# Test-FirewallProfiles

Custom compliance check: is Windows Firewall on for the domain, private and public profiles?

## What it returns

| Setting | Type | Compliant when |
|---|---|---|
| `DomainEnabled` | Boolean | true |
| `PrivateEnabled` | Boolean | true |
| `PublicEnabled` | Boolean | true |

The script reads the active policy store, so a profile turned on through Intune, GPO or locally all count.

## Setup in Intune

1. Devices > Compliance > Scripts: add `Test-FirewallProfiles.ps1` as a Windows discovery script (run as system, 64-bit).
2. In a Windows compliance policy, under Custom Compliance, pick the script and upload `Test-FirewallProfiles.json`.

Not using domain networks? Remove the `DomainEnabled` rule from the JSON.
