using OutlookHealthCheck.Models;

namespace OutlookHealthCheck.Diagnostics;

public sealed class DiagnosticApplicabilityEngine
{
    public bool Applies(DiagnosticDefinition definition, OutlookClientEnvironment environment, UserSymptom symptom)
    {
        var clientOk = definition.ApplicableClients.Count == 0 || definition.ApplicableClients.Contains(environment.ActiveClient) || definition.ApplicableClients.Contains(OutlookClientKind.Unknown);
        var symptomOk = definition.ApplicableSymptoms.Count == 0 || definition.ApplicableSymptoms.Contains(symptom) || definition.ApplicableSymptoms.Contains(UserSymptom.Unknown);
        return clientOk && symptomOk;
    }
}

public sealed class TimeSyncParser
{
    public DiagnosticResult Parse(string raw)
    {
        var leap = Find(raw, "Leap Indicator");
        var stratumText = Find(raw, "Stratum");
        var source = Find(raw, "Source");
        var lastSync = Find(raw, "Last Successful Sync Time");
        var status = DiagnosticStatus.Unknown;
        var finding = "Time synchronization state could not be determined from w32tm output.";
        var recommendation = "Review Windows Time configuration manually if authentication or Kerberos symptoms are present.";
        if (int.TryParse(stratumText?.Split(' ').FirstOrDefault(), out var stratum) && leap is not null)
        {
            if (leap.StartsWith("0", StringComparison.OrdinalIgnoreCase) && stratum > 0)
            {
                status = DiagnosticStatus.Pass;
                finding = "Windows Time reports a synchronized state.";
                recommendation = string.Empty;
            }
            else if (leap.StartsWith("3", StringComparison.OrdinalIgnoreCase) || stratum == 0)
            {
                status = DiagnosticStatus.Warning;
                finding = "Windows Time service may be running, but the clock is not synchronized.";
                recommendation = "Validate time source and sync state; authentication can fail when device time is unreliable.";
            }
        }
        return new DiagnosticResult
        {
            Id = "windows.time.sync",
            Name = "Windows Time Synchronization",
            Category = "System",
            Status = status,
            Severity = ToSeverity(status),
            Finding = finding,
            Recommendation = recommendation,
            RawResult = raw,
            Interpretation = $"Leap={leap ?? "unknown"}; Stratum={stratumText ?? "unknown"}; Source={source ?? "unknown"}; LastSync={lastSync ?? "unknown"}"
        };
    }

    private static string? Find(string raw, string name) => raw.Split('\n').Select(l => l.Trim()).FirstOrDefault(l => l.StartsWith(name, StringComparison.OrdinalIgnoreCase))?.Split(':', 2).ElementAtOrDefault(1)?.Trim();
    private static DiagnosticSeverity ToSeverity(DiagnosticStatus status) => Enum.Parse<DiagnosticSeverity>(status.ToString());
}
