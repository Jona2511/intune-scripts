# Autopilot

Scripts for Windows Autopilot. They use Microsoft Graph, so the notes in [Graph](../Graph) about permissions and sign-in apply here too.

| Script | What it does |
|---|---|
| [Invoke-AutopilotHashUpload](Invoke-AutopilotHashUpload) | Collects the hardware hash and uploads it to Intune, optionally with a group tag |
| [Set-AutopilotGroupTag](Set-AutopilotGroupTag) | Changes group tags in bulk from a CSV of serial numbers |
| [Get-AutopilotDevicesWithoutProfile](Get-AutopilotDevicesWithoutProfile) | Reports Autopilot devices without a successfully assigned deployment profile |
