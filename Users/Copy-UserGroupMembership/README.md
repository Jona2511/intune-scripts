# Copy-UserGroupMembership

Gives a user the same group memberships as a colleague. Handy when someone changes jobs or joins a team.

Only groups the source is a direct member of are copied. Dynamic groups, mail-enabled security groups, distribution lists and anything matching `-ExcludeGroups` are skipped, with the reason in the log. With `-RemoveOther` the target is also taken out of groups the source isn't in. The CSV lists every change and its result.

## Examples

```powershell
.\Copy-UserGroupMembership.ps1 -TenantId contoso.onmicrosoft.com -SourceUser k.bakker@contoso.com -TargetUser j.jansen@contoso.com -WhatIf
```

```powershell
.\Copy-UserGroupMembership.ps1 -TenantId $tid -SourceUser $colleague -TargetUser $user -RemoveOther -ExcludeGroups 'SG-Admins*'
```

## Permissions

User.Read.All and GroupMember.ReadWrite.All.
