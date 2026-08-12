# Diagnostics

Implemented diagnostic areas:

- System environment: computer, user, domain, OS, architecture, PowerShell/.NET runtime, drive free space, uptime, timezone.
- Outlook client environment: Classic installed/running, New Outlook installed/running, Office build/install type, active client classification.
- Outlook process: running/not running, responding state, PID/memory/path evidence where accessible.
- Profile inventory: multi-path HKCU discovery from Office and Windows Messaging Subsystem profile stores.
- Default profile validation: reconciles configured default profile against the discovered inventory.
- OST/PST/NST inventory: file presence, size, last write, accessibility, recent activity.
- Windows Time: parses `w32tm /query /status` for Leap Indicator, Stratum, Source, and Last Successful Sync Time.
- Microsoft 365 connectivity: endpoint-specific DNS and TCP 443 tests, plus WinHTTP proxy state.
- Event logs: recent Outlook crash/hang evidence from Application log for full checks.
- Windows services: collects state for ClickToRunSvc, VaultSvc, TokenBroker, CryptSvc, and BITS without treating manual stopped services as failures.

False-positive prevention:

- New Outlook installed is not treated as active unless running evidence exists.
- Missing profile registry evidence alone becomes `Unknown`, not `Critical`.
- `DiagnosticConsistencyEngine` detects profile contradictions such as empty registry discovery plus active Outlook/data-file/default-profile signals.
- Unknown and NotApplicable statuses do not reduce health score.
- Root-cause confidence is lowered when contradictions exist.
