using OutlookHealthCheck.Models;

namespace OutlookHealthCheck.Correlation;

public sealed class DiagnosticConsistencyEngine
{
    public List<Contradiction> FindContradictions(DiagnosticReport report)
    {
        var contradictions = new List<Contradiction>();
        var noProfileCritical = report.Diagnostics.Any(d => d.Id == "outlook.profiles" && d.Status == DiagnosticStatus.Critical);
        var activeClassicSignals = new List<string>();
        if (report.Outlook.ClassicRunning) activeClassicSignals.Add("OUTLOOK.EXE is running from the Office/Classic Outlook process family.");
        if (report.DataFiles.Any(f => f.RecentlyModified)) activeClassicSignals.Add("At least one Outlook data file was modified recently.");
        if (report.Profiles.DefaultProfileConfigured) activeClassicSignals.Add($"DefaultProfile is configured as '{report.Profiles.ConfiguredDefaultProfile}'.");
        if (report.Outlook.ActiveClient == OutlookClientKind.Classic) activeClassicSignals.Add("Active client detection points to Classic Outlook.");

        if ((noProfileCritical || report.Profiles.Profiles.Count == 0) && activeClassicSignals.Count >= 2)
        {
            contradictions.Add(new Contradiction
            {
                Title = "Profile diagnostics produced conflicting evidence",
                Conclusion = "Profile health cannot be reliably determined from registry evidence alone.",
                SupportingEvidence = ["Profile registry discovery did not find a supported profile."],
                ContradictingEvidence = activeClassicSignals
            });
        }

        if (report.Outlook.NewOutlookInstalled && report.Outlook.ClassicRunning)
        {
            contradictions.Add(new Contradiction
            {
                Title = "New Outlook installed but Classic Outlook active",
                Conclusion = "New Outlook package presence must not be treated as the active client.",
                SupportingEvidence = ["New Outlook installed indicator is present."],
                ContradictingEvidence = ["OUTLOOK.EXE is running, which is Classic Outlook evidence."]
            });
        }

        if (report.Profiles.DefaultProfileConfigured && report.Profiles.Profiles.Count == 0)
        {
            contradictions.Add(new Contradiction
            {
                Title = "Default profile configured but profile inventory empty",
                Conclusion = "Default profile validation is inconclusive until profile registry structures are manually verified.",
                SupportingEvidence = [$"DefaultProfile registry value: {report.Profiles.ConfiguredDefaultProfile}"],
                ContradictingEvidence = ["No profile key was discovered in supported registry locations."]
            });
        }
        return contradictions;
    }
}

public sealed class HealthScoringService
{
    private static readonly Dictionary<string, int> Weights = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Authentication"] = 30,
        ["Microsoft 365"] = 25,
        ["Outlook Profile"] = 25,
        ["Outlook"] = 25,
        ["Office"] = 15,
        ["Network"] = 20,
        ["Event Logs"] = 20,
        ["Add-ins"] = 5,
        ["System"] = 5,
        ["OST/PST"] = 15
    };

    public HealthScoreResult Score(IEnumerable<DiagnosticResult> diagnostics, IReadOnlyList<Contradiction> contradictions, IReadOnlyList<RootCauseHypothesis> rootCauses)
    {
        var list = diagnostics.ToList();
        var categories = list.GroupBy(d => d.Category);
        double penalty = 0;
        double totalWeight = 0;
        foreach (var group in categories)
        {
            var weight = Weights.GetValueOrDefault(group.Key, 10);
            var applicable = group.Where(d => d.Status is not DiagnosticStatus.Unknown and not DiagnosticStatus.NotApplicable and not DiagnosticStatus.Error).ToList();
            if (applicable.Count == 0) continue;
            totalWeight += weight;
            var categoryPenalty = applicable.Max(d => d.Status switch
            {
                DiagnosticStatus.Critical => 1.0,
                DiagnosticStatus.Warning => 0.45,
                _ => 0
            });
            penalty += weight * categoryPenalty;
        }

        var score = totalWeight <= 0 ? 100 : Math.Clamp((int)Math.Round(100 - (penalty / totalWeight * 100)), 0, 100);
        var confidence = contradictions.Count > 0 ? ConfidenceLevel.Low : list.Count(d => d.Status is DiagnosticStatus.Pass or DiagnosticStatus.Info or DiagnosticStatus.Warning or DiagnosticStatus.Critical) >= 6 ? ConfidenceLevel.Medium : ConfidenceLevel.Low;
        var rootConfidence = rootCauses.Select(r => r.Confidence).DefaultIfEmpty(ConfidenceLevel.Unknown).Max();
        return new HealthScoreResult
        {
            HealthScore = score,
            DiagnosticConfidence = confidence,
            RootCauseConfidence = contradictions.Count > 0 && rootConfidence == ConfidenceLevel.High ? ConfidenceLevel.Medium : rootConfidence,
            Counts = Enum.GetValues<DiagnosticStatus>().ToDictionary(s => s.ToString(), s => list.Count(d => d.Status == s))
        };
    }
}

public sealed class RootCauseEngine
{
    public List<RootCauseHypothesis> Analyze(DiagnosticReport report)
    {
        if (report.Contradictions.Count > 0)
        {
            return
            [
                new RootCauseHypothesis
                {
                    Category = "Conflicting Evidence",
                    Hypothesis = "Available diagnostics conflict; root cause cannot be assigned with high confidence.",
                    Confidence = ConfidenceLevel.Low,
                    SupportingEvidence = report.Contradictions.Select(c => c.Title).ToList(),
                    ContradictingEvidence = report.Contradictions.SelectMany(c => c.ContradictingEvidence).ToList(),
                    AffectedComponents = ["Outlook", "Profile", "Local evidence model"],
                    RecommendedNextStep = "Validate the conflicting evidence manually before recreating profiles or deleting data files."
                }
            ];
        }

        var findings = report.Diagnostics.Where(d => d.Status is DiagnosticStatus.Warning or DiagnosticStatus.Critical).ToList();
        var hypotheses = new List<RootCauseHypothesis>();
        AddIfAny(hypotheses, findings, "Authentication", "Office identity or authentication state requires attention", ["Authentication", "Office", "WAM"], "Validate Office identity, WAM, Credential Manager indicators, and activation state before profile remediation.");
        AddIfAny(hypotheses, findings, "Network", "Microsoft 365 connectivity path may be impaired", ["DNS", "TCP 443", "TLS", "Proxy"], "Separate DNS, TCP 443, HTTPS, and proxy results to identify the failed layer.");
        AddIfAny(hypotheses, findings, "Outlook Profile", "Local Outlook profile/data configuration may require review", ["Profile", "OST/PST", "Classic Outlook"], "Verify Control Panel > Mail profile inventory and default profile before rebuilding anything.");
        AddIfAny(hypotheses, findings, "Event Logs", "Recent Outlook crash or hang evidence was found", ["Outlook", "Add-ins", "Office build"], "Review crash frequency, faulting module, and add-in state; test safe mode if crashes repeat.");
        return hypotheses.OrderByDescending(h => h.Confidence).ToList();
    }

    private static void AddIfAny(List<RootCauseHypothesis> output, List<DiagnosticResult> findings, string category, string hypothesis, List<string> components, string nextStep)
    {
        var items = findings.Where(f => f.Category.Equals(category, StringComparison.OrdinalIgnoreCase) || (category == "Outlook Profile" && f.Category is "OST/PST" or "Outlook Profile")).ToList();
        if (items.Count == 0) return;
        var independent = items.Select(i => i.Id.Split('.').FirstOrDefault() ?? i.Category).Distinct().Count();
        var critical = items.Any(i => i.Status == DiagnosticStatus.Critical);
        output.Add(new RootCauseHypothesis
        {
            Category = category,
            Hypothesis = hypothesis,
            Confidence = critical && independent >= 2 ? ConfidenceLevel.Medium : ConfidenceLevel.Low,
            SupportingEvidence = items.Select(i => $"{i.Name}: {i.Finding}").ToList(),
            ContradictingEvidence = [],
            AffectedComponents = components,
            RecommendedNextStep = nextStep
        });
    }
}

public sealed class EngineerSummaryGenerator
{
    public string Generate(DiagnosticReport report)
    {
        var lines = new List<string> { $"The selected symptom is {report.Symptom}." };
        lines.Add(report.Outlook.ActiveClient switch
        {
            OutlookClientKind.Classic => "Classic Outlook is currently the active client based on OUTLOOK.EXE process evidence.",
            OutlookClientKind.NewOutlook => "New Outlook appears to be the active client based on running process evidence.",
            OutlookClientKind.Both => "Both Classic Outlook and New Outlook process evidence is present.",
            OutlookClientKind.None => "No Outlook client was detected as installed or running.",
            _ => "The active Outlook client could not be determined reliably."
        });
        if (report.Contradictions.Count > 0)
        {
            lines.Add("Conflicting evidence was detected; do not treat any single registry absence as proof of a broken profile.");
            lines.Add("Recommended next step: validate the profile inventory and active client manually before destructive remediation.");
            return string.Join(Environment.NewLine, lines);
        }
        var top = report.RootCauses.FirstOrDefault();
        if (top is null) lines.Add("No critical issues or weighted warning patterns were detected by the automated diagnostics.");
        else lines.Add($"Evidence currently points toward: {top.Hypothesis} (confidence: {top.Confidence}).");
        lines.Add($"Recommended next step: {top?.RecommendedNextStep ?? "Gather the user symptom timeline and review the detailed evidence."}");
        return string.Join(Environment.NewLine, lines);
    }
}
