# Outlook Health Check
### Microsoft Outlook Diagnostic & Troubleshooting Assistant

A read-only, diagnostic-first PowerShell + WPF application for IT/Messaging
support engineers troubleshooting Outlook on Windows in a Microsoft 365 /
Exchange Online environment. Runs entirely in the interactive signed-in
user's own context — no elevation, no Exchange Online admin connection.

## Running it

1. Unzip the `OutlookHealthCheck` folder anywhere (e.g. Desktop).
2. Windows may mark the downloaded files as "blocked." Unblock them first:
   ```powershell
   Get-ChildItem -Path .\OutlookHealthCheck -Recurse | Unblock-File
   ```
3. Launch it:
   ```powershell
   cd .\OutlookHealthCheck
   .\OutlookHealthCheck.ps1
   ```
   If your execution policy blocks the script, run it this way instead
   (this only affects the current PowerShell session, not the system-wide
   policy — the safest way to run a script you've reviewed):
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\OutlookHealthCheck.ps1
   ```
4. Choose **Quick Check** (~20 seconds, the essentials) or **Full Check**
   (comprehensive — system, Outlook, Office, identity, network, event logs).
5. Click **RUN HEALTH CHECK**. When it finishes, the HTML report opens
   automatically in your default browser and is also saved under `Reports\`.

## What it does

- Collects diagnostics across System, Outlook, Office, Authentication,
  Network, and Event Logs (~30 individual checks).
- Every check returns a structured result: Category, Check, Status/Severity,
  Value, Expected, Finding, Recommendation, and the exact PowerShell command
  used — all visible in the report's expandable "Technical Details."
- A correlation engine classifies non-passing findings into root-cause
  buckets (Local Outlook, Office Installation, Windows, Authentication,
  Network) and assigns a confidence level (HIGH/MEDIUM/LOW) based on the
  pattern of what's healthy vs. abnormal.
- Produces a single self-contained HTML file plus a matching JSON export.

## What it will never do

This tool is **strictly read-only**. It never:
- Modifies or deletes registry keys
- Deletes, recreates, or repairs OST/PST files
- Disables/enables/uninstalls add-ins
- Removes or modifies Credential Manager entries
- Repairs Office, restarts services, or changes network/proxy settings
- Attempts Exchange Online / admin-level connectivity

## Privacy

Passwords, tokens, cookies, license keys, and other secrets are never
collected. Where credential-related registry/Credential Manager entries are
inspected, only entry *names* are shown — never values. Report text passes
through a redaction filter before being written to the log or HTML/JSON.

## Project layout

```text
OutlookHealthCheck/
├── OutlookHealthCheck.ps1
├── Modules/
│   ├── Common.ps1
│   ├── SystemChecks.ps1
│   ├── OutlookChecks.ps1
│   ├── OfficeChecks.ps1
│   ├── IdentityChecks.ps1
│   ├── NetworkChecks.ps1
│   ├── EventLogChecks.ps1
│   ├── CorrelationEngine.ps1
│   └── ReportGenerator.ps1
├── Reports/
└── Logs/
```

## Extending it

Every check is an independent function returning a `New-CheckResult` object.
Add new checks to the appropriate module and its aggregator so the results
flow through scoring, correlation, and reporting.
