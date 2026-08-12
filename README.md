# Outlook Health Check

Windows Outlook Troubleshooting & Diagnostic Assistant for Messaging/IT support engineers.

This repository contains two implementations:

- `OutlookHealthCheck.ps1`: original PowerShell/WPF script retained for reference.
- `src/OutlookHealthCheck`: .NET 8 WPF application with typed evidence, applicability, consistency, root-cause, reporting, and testable diagnostic logic.

## Build

```powershell
dotnet restore .\OutlookHealthCheck.slnx
dotnet build .\OutlookHealthCheck.slnx
```

## Run UI

```powershell
dotnet run --project .\src\OutlookHealthCheck\OutlookHealthCheck.csproj
```

## Run Diagnostics From CLI

```powershell
dotnet run --project .\src\OutlookHealthCheck\OutlookHealthCheck.csproj -- --quick --symptom=Unknown
dotnet run --project .\src\OutlookHealthCheck\OutlookHealthCheck.csproj -- --full --symptom=PasswordPrompts
```

Reports are written to:

```text
src\OutlookHealthCheck\bin\Debug\net8.0-windows\Reports
```

Logs are written to:

```text
src\OutlookHealthCheck\bin\Debug\net8.0-windows\Logs
```

## Test

The test project is a zero-NuGet regression runner so it can execute in locked-down environments:

```powershell
dotnet run --project .\tests\OutlookHealthCheck.Tests\OutlookHealthCheck.Tests.csproj
```

## Design Rule

The assistant never marks a state broken unless collected evidence supports that conclusion. Incomplete evidence is `Unknown`; irrelevant diagnostics are `NotApplicable`; conflicting evidence is surfaced explicitly.
