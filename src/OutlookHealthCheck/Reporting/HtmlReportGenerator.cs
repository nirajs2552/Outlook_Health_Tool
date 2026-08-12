using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookHealthCheck.Models;
using OutlookHealthCheck.Services;

namespace OutlookHealthCheck.Reporting;

public sealed class HtmlReportGenerator(IRedactor redactor)
{
    public string ReportDirectory { get; } = Path.Combine(AppContext.BaseDirectory, "Reports");

    public void Write(DiagnosticReport report)
    {
        Directory.CreateDirectory(ReportDirectory);
        var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var html = Path.Combine(ReportDirectory, $"Outlook-Health-Report-{stamp}.html");
        var json = Path.Combine(ReportDirectory, $"Outlook-Health-Report-{stamp}.json");
        File.WriteAllText(html, BuildHtml(report), Encoding.UTF8);
        File.WriteAllText(json, redactor.Redact(JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true, Converters = { new JsonStringEnumConverter() } })), Encoding.UTF8);
        report.HtmlReportPath = html;
        report.JsonReportPath = json;
    }

    private static string BuildHtml(DiagnosticReport report)
    {
        var e = HtmlEncoder.Default;
        var critical = report.Diagnostics.Where(d => d.Status == DiagnosticStatus.Critical).ToList();
        var warnings = report.Diagnostics.Where(d => d.Status == DiagnosticStatus.Warning).ToList();
        var sb = new StringBuilder();
        sb.Append("""
<!doctype html><html><head><meta charset="utf-8"><title>Outlook Health Report</title><style>
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f5f7fa;color:#172033}header{background:#1f3a5f;color:white;padding:24px 34px}main{padding:24px 34px}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.panel{background:white;border:1px solid #d9e1ea;border-radius:6px;padding:16px;margin-bottom:16px}.badge{display:inline-block;padding:3px 8px;border-radius:4px;font-weight:700;font-size:12px}.Critical{background:#b42318;color:white}.Warning{background:#f59e0b;color:#111827}.Pass{background:#15803d;color:white}.Info,.Unknown,.NotApplicable{background:#e5e7eb;color:#111827}.critical-card{border-left:6px solid #b42318}.mono{font-family:Consolas,monospace;white-space:pre-wrap;background:#f8fafc;padding:10px;border:1px solid #e5e7eb;border-radius:4px;overflow:auto}table{border-collapse:collapse;width:100%}td,th{border-bottom:1px solid #e5e7eb;text-align:left;padding:8px}h2{margin-top:0}.score{font-size:42px;font-weight:700}
</style></head><body>
""");
        sb.Append($"<header><h1>Outlook Health Check</h1><div>Microsoft Outlook Diagnostic Assistant</div></header><main>");
        Section(sb, "1. Executive Summary", $"<div class='mono'>{e.Encode(report.EngineerSummary)}</div>");
        Section(sb, "2. User Symptom", e.Encode(report.Symptom.ToString()));
        Section(sb, "3. Overall Health", $"<div class='score'>{report.Health.HealthScore} / 100</div><p>Diagnostic confidence: {report.Health.DiagnosticConfidence}<br>Root cause confidence: {report.Health.RootCauseConfidence}</p>");
        Section(sb, "4. Diagnostic Confidence", Counts(report));
        Section(sb, "5. Critical Findings", critical.Count == 0 ? "<p>No critical findings were produced.</p>" : string.Join("", critical.Select(d => FindingCard(d, true))));
        Section(sb, "6. Likely Root Causes", report.RootCauses.Count == 0 ? "<p>No root-cause hypothesis met evidence threshold.</p>" : string.Join("", report.RootCauses.Select(RootCauseCard)));
        Section(sb, "7. Contradictory Evidence", report.Contradictions.Count == 0 ? "<p>No contradictions detected.</p>" : string.Join("", report.Contradictions.Select(ContradictionCard)));
        Section(sb, "8. Recommended Next Steps", report.RootCauses.FirstOrDefault()?.RecommendedNextStep ?? "Review the detailed evidence and symptom timeline.");
        Section(sb, "9. Environment", $"<table><tr><th>Computer</th><td>{e.Encode(report.Environment.ComputerName)}</td></tr><tr><th>User</th><td>{e.Encode(report.Environment.Domain)}\\{e.Encode(report.Environment.UserName)}</td></tr><tr><th>OS</th><td>{e.Encode(report.Environment.OS)}</td></tr><tr><th>PowerShell</th><td>{e.Encode(report.Environment.PowerShellVersion)}</td></tr></table>");
        Section(sb, "10. Outlook", $"<table><tr><th>Classic Installed</th><td>{report.Outlook.ClassicInstalled}</td></tr><tr><th>Classic Running</th><td>{report.Outlook.ClassicRunning}</td></tr><tr><th>New Installed</th><td>{report.Outlook.NewOutlookInstalled}</td></tr><tr><th>New Running</th><td>{report.Outlook.NewOutlookRunning}</td></tr><tr><th>Active Client</th><td>{report.Outlook.ActiveClient}</td></tr></table>");
        Section(sb, "11. Profiles", Profiles(report));
        Section(sb, "12. OST/PST", DataFiles(report));
        foreach (var groupName in new[] { "Office", "Identity", "Network", "DNS", "Proxy", "Add-ins", "Event Logs", "Windows Services", "Updates", "Technical Evidence" })
        {
            var rows = report.Diagnostics.Where(d => groupName switch { "DNS" => d.Id.Contains("network.dns"), "Proxy" => d.Id.Contains("network.proxy"), "Windows Services" => d.Id.Contains("windows.services"), "Technical Evidence" => true, _ => d.Category.Equals(groupName, StringComparison.OrdinalIgnoreCase) }).ToList();
            Section(sb, groupName, rows.Count == 0 ? "<p>No diagnostics in this section.</p>" : string.Join("", rows.Select(d => FindingCard(d, false))));
        }
        sb.Append("</main></body></html>");
        return sb.ToString();
    }

    private static void Section(StringBuilder sb, string title, string body) => sb.Append($"<section class='panel'><h2>{title}</h2>{body}</section>");
    private static string Counts(DiagnosticReport r) => "<div class='grid'>" + string.Join("", r.Health.Counts.Select(kv => $"<div class='panel'><b>{kv.Key}</b><div>{kv.Value}</div></div>")) + "</div>";
    private static string FindingCard(DiagnosticResult d, bool critical) => $"<div class='panel {(critical ? "critical-card" : "")}'><span class='badge {d.Status}'>{d.Status}</span><h3>{HtmlEncoder.Default.Encode(d.Name)}</h3><p><b>Finding:</b> {HtmlEncoder.Default.Encode(d.Finding)}</p><p><b>Interpretation:</b> {HtmlEncoder.Default.Encode(d.Interpretation)}</p><p><b>Recommendation:</b> {HtmlEncoder.Default.Encode(d.Recommendation)}</p><details><summary>Raw Evidence</summary><div class='mono'>{HtmlEncoder.Default.Encode(d.RawResult)}</div></details></div>";
    private static string RootCauseCard(RootCauseHypothesis h) => $"<div class='panel'><h3>{HtmlEncoder.Default.Encode(h.Hypothesis)}</h3><p><b>Category:</b> {h.Category}<br><b>Confidence:</b> {h.Confidence}</p><p><b>Supporting:</b><br>{string.Join("<br>", h.SupportingEvidence.Select(HtmlEncoder.Default.Encode))}</p><p><b>Contradicting:</b><br>{string.Join("<br>", h.ContradictingEvidence.Select(HtmlEncoder.Default.Encode))}</p><p><b>Next step:</b> {HtmlEncoder.Default.Encode(h.RecommendedNextStep)}</p></div>";
    private static string ContradictionCard(Contradiction c) => $"<div class='panel critical-card'><h3>CONFLICTING / INCONCLUSIVE EVIDENCE: {HtmlEncoder.Default.Encode(c.Title)}</h3><p>{HtmlEncoder.Default.Encode(c.Conclusion)}</p><p><b>Supporting:</b><br>{string.Join("<br>", c.SupportingEvidence.Select(HtmlEncoder.Default.Encode))}</p><p><b>But:</b><br>{string.Join("<br>", c.ContradictingEvidence.Select(HtmlEncoder.Default.Encode))}</p></div>";
    private static string Profiles(DiagnosticReport r) => r.Profiles.Profiles.Count == 0 ? $"<p>Profiles discovered: 0</p><p>{HtmlEncoder.Default.Encode(r.Profiles.EvidenceSummary)}</p>" : "<table><tr><th>Profile</th><th>Default</th><th>Source</th><th>Evidence</th></tr>" + string.Join("", r.Profiles.Profiles.Select(p => $"<tr><td>{HtmlEncoder.Default.Encode(p.ProfileName)}</td><td>{(p.IsDefault?.ToString() ?? "UNKNOWN")}</td><td>{HtmlEncoder.Default.Encode(p.Source)}</td><td>{HtmlEncoder.Default.Encode(p.ProfileEvidence)}</td></tr>")) + "</table>";
    private static string DataFiles(DiagnosticReport r) => r.DataFiles.Count == 0 ? "<p>No data files discovered in the default local Outlook data folder.</p>" : "<table><tr><th>Path</th><th>Size GB</th><th>Last Write</th><th>Accessible</th></tr>" + string.Join("", r.DataFiles.Select(f => $"<tr><td>{HtmlEncoder.Default.Encode(Path.GetFileName(f.Path))}</td><td>{f.SizeGb}</td><td>{f.LastWriteTime}</td><td>{f.Accessible}</td></tr>")) + "</table>";
}


