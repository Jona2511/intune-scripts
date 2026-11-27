# Invoke-LocalAdminCheck

Detects unexpected members of the local Administrators group and reports or removes them.

## Description

```text
Reads the members of the local Administrators group (found by SID S-1-5-32-544, so it
also works on Dutch Windows, where the group is called "Administrators"). Uses ADSI
instead of Get-LocalGroupMember, because that cmdlet fails on Entra ID members and
orphaned SIDs.

Each member is compared against an allow list. Entries can be:
  - account names:  'CONTOSO\IT-Admins', 'LAPSAdmin', '*\Domain Admins'
  - SIDs:           'S-1-12-1-1234-...'
  - SID wildcards:  'S-1-5-21-*-512'
Always allowed (built in):
  - the built-in local Administrator account (local SID ending in -500)
  - Domain Admins (domain SID ending in -512) on domain-joined devices

Entra ID role and group members (e.g. Global Administrator, Microsoft Entra Joined
Device Local Administrator) appear as unresolvable S-1-12-1-... SIDs. Allow them with
-AllowedEntraObjectIds <object ID of the role/group>; the script converts the GUID to the SID.

Default mode is report only. With -Remove, unexpected members are removed (supports -WhatIf).
Safety: no removal happens if afterwards no enabled administrator would remain.
```

## Parameters

- `-AllowedMembers`: Allowed names/SIDs (wildcards allowed).
- `-AllowListPath`: Text file with extra allowed entries, one per line ('#' = comment).
- `-AllowedEntraObjectIds`: Object IDs (GUIDs) of Entra ID roles or groups that may be members.
- `-Remove`: Removes unexpected members. Without this switch the script only reports.

## Examples

```powershell
.\Invoke-LocalAdminCheck.ps1 -AllowedMembers 'CONTOSO\Workstation Admins'
```

```powershell
.\Invoke-LocalAdminCheck.ps1 -AllowedEntraObjectIds '9a1b...','4c2d...' -Remove -WhatIf
```

## Notes

```text
Exit codes: 0 = no unexpected members (afterwards), 1 = unexpected members present /
removal failed or blocked, 2 = fatal error.
Log: %ProgramData%\LocalAdminCheck\
```
