# New-M365User

Creates a Microsoft 365 user with licence, manager and groups in one go.

The script looks up the manager, licence and groups before it creates anything, so a typo stops it before the account exists. It also checks that there is a free licence. The temporary password appears once on screen and never in the log. The user has to change it at first sign-in.

Use `-CopyGroupsFrom` to give the new user the same groups as a colleague in the same role. Dynamic groups and mail-enabled groups are skipped, because Graph can't change those.

## Example

```powershell
.\New-M365User.ps1 -TenantId contoso.onmicrosoft.com -UserPrincipalName j.jansen@contoso.com `
    -GivenName Jan -Surname Jansen -JobTitle Sales -Manager p.devries@contoso.com `
    -LicenseSku SPB -CopyGroupsFrom k.bakker@contoso.com -WhatIf
```

Leave out `-LicenseSku` if the tenant uses group-based licensing, and add the licence group to `-Groups` instead. [Get-LicenseOverview](../../Graph/Get-LicenseOverview) shows the SkuPartNumbers in a tenant (`SPB` is Microsoft 365 Business Premium).

## Permissions

User.ReadWrite.All, GroupMember.ReadWrite.All and Organization.Read.All, as application permissions or delegated (signed in as User Administrator or higher).
