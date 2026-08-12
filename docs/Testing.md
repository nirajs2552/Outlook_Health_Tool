# Testing

The test project is intentionally package-free and uses a small local assertion runner. It covers the required regression classes:

- Profile discovery: one/two/multiple/no profiles, valid/invalid defaults, New Outlook and Classic Outlook client combinations.
- Evidence handling: OST exists/missing/recent, Outlook running/not running/responding/not responding.
- Contradiction handling: no profile plus active Outlook/data-file signals, default profile with empty inventory, New Outlook installed plus Classic active.
- Time parsing: synchronized and service-running-but-not-synchronized output.
- Network model: DNS pass/fail, TCP 443 failure, proxy detected.
- Event model: no crash, one crash, repeated crash, faulting module evidence.

Run:

```powershell
dotnet run --project .\tests\OutlookHealthCheck.Tests\OutlookHealthCheck.Tests.csproj
```

Every future diagnostic bug should be represented as a regression test against the evidence/interpreter layer, not only a UI text change.
