# Platform scripts

Scripts for Devices > Scripts and remediations > Platform scripts. Intune runs these once per device, and again only when the script changes.

Use these for one-time configuration. Anything that has to stay correct over time is better off as a remediation.

| Script | What it does |
|---|---|
| [Invoke-BloatwareRemoval](Invoke-BloatwareRemoval) | Removes provisioned Windows apps that are not on an allowlist |
| [Invoke-RegionalSettings](Invoke-RegionalSettings) | Sets the time zone, Dutch regional formats and home location |
| [Invoke-StartTaskbarCleanup](Invoke-StartTaskbarCleanup) | Unpins the default apps from the taskbar and Start menu on new Windows 11 devices |
