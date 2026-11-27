# Invoke-UserOffboarding

Offboards a Microsoft 365 user.

Blocking sign-in and revoking sessions always happen. The rest is opt-in:

| Switch | Does |
|---|---|
| `-ResetPassword` | Sets a random password that nobody sees |
| `-DisableDevices` | Disables the user's devices in Entra ID |
| `-RemoveGroups` | Removes the user from groups (dynamic and mail-enabled groups are left alone) |
| `-RemoveLicenses` | Removes directly assigned licences; group-based ones go with the groups |

Before it changes anything, the script exports the user's groups, licences and devices to a CSV, so you can put them back.

Convert the mailbox to a shared mailbox **before** you use `-RemoveLicenses`. Exchange deletes an unlicensed mailbox after 30 days. The mailbox itself, forwarding and giving the manager access to OneDrive are not part of this script.

## Example

```powershell
.\Invoke-UserOffboarding.ps1 -TenantId contoso.onmicrosoft.com -User j.jansen@contoso.com -RemoveGroups -RemoveLicenses -WhatIf
```

## Permissions

User.ReadWrite.All and GroupMember.ReadWrite.All. For `-DisableDevices` also Device.ReadWrite.All. For `-ResetPassword` with app-only sign-in, User-PasswordProfile.ReadWrite.All.
