# Get-StaleGuestUsers

Reports guest accounts that are no longer used. Report only, nothing is removed.

A guest is listed when the invitation was never accepted, when the last sign-in is older than `-Days` (default 90), or when the guest never signed in at all. The report includes the guest's domain and the groups and Teams they are still in, so you can ask the right partner or team owner before you clean up. `-SkipGroups` makes it faster in tenants with many guests.

## Example

```powershell
.\Get-StaleGuestUsers.ps1 -TenantId contoso.onmicrosoft.com -Days 180
```

## Permissions

User.Read.All, AuditLog.Read.All, GroupMember.Read.All and Organization.Read.All. Sign-in dates need Entra ID P1 or P2.
