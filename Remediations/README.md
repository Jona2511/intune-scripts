# Remediations

Detection and remediation scripts for Intune Remediations (Devices > Scripts and remediations).

Each pair sits in its own folder:

```
Remediations/
  <Name>/
    Detect-<Name>.ps1
    Remediate-<Name>.ps1
```

The detection script exits with 1 when the device needs fixing, so Intune runs the remediation, and with 0 when it doesn't. Both scripts write one short line to the output, which Intune shows in the device status.

Start from the templates in [Templates](../Templates).
