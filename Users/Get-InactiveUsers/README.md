# Get-InactiveUsers

Reports member accounts that haven't signed in for more than X days (default 90). Report only, nothing is changed.

The script takes the most recent of the interactive, non-interactive and successful sign-in dates. Accounts that never signed in count once they are older than the cutoff. Each row shows the licences, and the log ends with a count of inactive users per licence, which shows quickly what you're paying for without it being used.

Disabled accounts are left out unless you add `-IncludeDisabled`. Guests have their own report: [Get-StaleGuestUsers](../Get-StaleGuestUsers).

## Example

```powershell
.\Get-InactiveUsers.ps1 -TenantId contoso.onmicrosoft.com -Days 60
```

## Permissions

User.Read.All, AuditLog.Read.All and Organization.Read.All. Sign-in dates need Entra ID P1 or P2, which Microsoft 365 Business Premium includes.
