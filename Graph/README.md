# Graph

Scripts that use the [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/powershell/microsoftgraph/overview) for reporting and bulk changes.

Each script lists the permissions it needs in its header and asks for them with `Connect-MgGraph -Scopes`. None of them store credentials. For unattended runs, use an app registration with a certificate instead of a client secret.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
