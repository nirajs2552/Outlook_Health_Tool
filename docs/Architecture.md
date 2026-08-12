# Architecture

The .NET implementation follows this flow:

```text
WPF Application
  -> DiagnosticController
  -> EnvironmentDetector
  -> Diagnostic Providers / Evidence Collectors
  -> Evidence Model
  -> Applicability / Interpretation
  -> DiagnosticConsistencyEngine
  -> RootCauseEngine
  -> WPF Dashboard + HTML/JSON Reports
```

Concerns are separated by folder:

- `UI`: WPF views and view model.
- `Diagnostics`: orchestration, applicability, parsers, diagnostic status production.
- `Evidence`: environment, profile, and local evidence collectors.
- `Correlation`: contradiction detection, health scoring, root-cause hypotheses, engineer summary.
- `Reporting`: self-contained HTML and structured JSON export.
- `PowerShell`: controlled PowerShell process execution with timeout, cancellation, output/error capture, and redaction.
- `Models`: strongly typed report, evidence, profile, status, confidence, and diagnostic models.
- `Services`: registry, process, logging, and redaction boundaries.

Normal diagnostics are read-only. The application does not delete OST/PST files, modify registry keys, remove credentials, change services, change proxy, or repair Office.
