# Compliance

Custom compliance settings for Windows. Each check has two files:

```
Compliance/
  <Name>/
    Discover-<Name>.ps1
    <Name>.json
```

The discovery script returns its results as one line of compressed JSON (`ConvertTo-Json -Compress`). The JSON file holds the rules Intune checks those results against, including the message users see when a device is not compliant.
