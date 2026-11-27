# Invoke-AutopilotHashUpload

Collects the Autopilot hardware hash and uploads it to Intune through Microsoft Graph,
authenticating as an app registration with a certificate. Sets the group tag directly.

## Description

```text
Built to run in OOBE (Shift+F10 -> powershell) or on an existing device, without any
PowerShell modules (no Microsoft.Graph, no WindowsAutopilotIntune, no MSAL):

  1. Reads the hardware hash (MDM_DevDetail_Ext01 in root/cimv2/mdm/dmmap), the serial
     number (Win32_BIOS) and, if present, the OEM product key.
  2. Optionally saves them to a CSV in the Intune import format (-OutputCsv).
  3. Builds a JWT client assertion signed with the certificate's private key (RS256,
     x5t header) and gets an app-only token from login.microsoftonline.com.
     Certificate source: -CertificateThumbprint (LocalMachine\My, then CurrentUser\My)
     or -PfxPath + -PfxPassword. Works with CNG and legacy CSP keys.
  4. Checks whether the serial number already exists in Autopilot. If it does, it reports
     it and, with -UpdateExisting, sets the group tag (and assigned user) through the
     updateDeviceProperties action instead of importing again.
  5. Imports the device (POST importedWindowsAutopilotDeviceIdentities) with groupTag,
     polls the import until 'complete' or 'error', and removes the import record.
  6. Unless -NoWait: polls windowsAutopilotDeviceIdentities until the device appears
     and reports the profile assignment status. With -WaitForProfile it keeps waiting
     until a deployment profile is assigned.

App registration requirements:
  - API permission (Application): DeviceManagementServiceConfig.ReadWrite.All, admin consent.
  - The certificate's public key (.cer) uploaded under "Certificates & secrets".
```

## Parameters

- `-TenantId`: Tenant ID (GUID) or primary domain, e.g. 'contoso.onmicrosoft.com'.
- `-ClientId`: Application (client) ID of the app registration.
- `-CertificateThumbprint`: Thumbprint of a certificate with private key in LocalMachine\My or CurrentUser\My.
- `-PfxPath`: Path to a .pfx file with the certificate and private key (e.g. on a USB stick).
- `-PfxPassword`: Password of the .pfx as SecureString.
- `-GroupTag`: Autopilot group tag (OrderID) to set.
- `-AssignedUser`: Optional UPN to assign to the device.
- `-UpdateExisting`: If the serial already exists in Autopilot: update its group tag/assigned user.
- `-OutputCsv`: Also save the hash to this CSV (Intune import format).
- `-TimeoutMinutes`: Maximum total waiting time for import and appearance (default 25).
- `-NoWait`: Do not wait for the device to appear in Autopilot after a successful import.
- `-WaitForProfile`: Keep waiting until a deployment profile is assigned (within -TimeoutMinutes).

## Examples

```powershell
# OOBE, certificate from USB stick
$pw = Read-Host -AsSecureString 'PFX password'
.\Invoke-AutopilotHashUpload.ps1 -TenantId contoso.onmicrosoft.com -ClientId <guid> `
    -PfxPath D:\autopilot.pfx -PfxPassword $pw -GroupTag 'NL-Standard' -WaitForProfile
```

```powershell
.\Invoke-AutopilotHashUpload.ps1 -TenantId <guid> -ClientId <guid> -CertificateThumbprint <thumb> `
    -GroupTag 'NL-Kiosk' -UpdateExisting -OutputCsv C:\Temp\hash.csv
```

## Notes

```text
Exit codes: 0 = imported / already present (and updated), 1 = import error, timeout or
device already registered without -UpdateExisting, 2 = fatal error.
Log: %ProgramData%\AutopilotHashUpload\
Treat a PFX with private key as a secret: anyone holding it can register devices
in your tenant. Prefer a dedicated app registration with only this permission.
```
