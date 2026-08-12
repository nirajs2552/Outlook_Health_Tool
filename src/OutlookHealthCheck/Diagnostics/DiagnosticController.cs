using System.Diagnostics;
using OutlookHealthCheck.Correlation;
using OutlookHealthCheck.Evidence;
using OutlookHealthCheck.Models;
using OutlookHealthCheck.PowerShell;
using OutlookHealthCheck.Reporting;
using OutlookHealthCheck.Services;

namespace OutlookHealthCheck.Diagnostics;

public sealed class DiagnosticController(
    OutlookEnvironmentDetector environmentDetector,
    OutlookProfileDiscovery profileDiscovery,
    LocalEvidenceCollector evidenceCollector,
    IProcessInspector processInspector,
    IPowerShellExecutor powerShell,
    DiagnosticConsistencyEngine consistencyEngine,
    RootCauseEngine rootCauseEngine,
    HealthScoringService scoringService,
    EngineerSummaryGenerator summaryGenerator,
    HtmlReportGenerator reportGenerator,
    DiagnosticLogger logger)
{
    public event Action<int, string>? ProgressChanged;

    public async Task<DiagnosticReport> RunAsync(UserSymptom symptom, bool full, CancellationToken cancellationToken)
    {
        logger.Start();
        var sw = Stopwatch.StartNew();
        var report = new DiagnosticReport { Symptom = symptom };
        var diagnostics = report.Diagnostics;

        Step(5, "Detecting Outlook client environment...");
        report.Environment = evidenceCollector.CollectEnvironment();
        report.Outlook = environmentDetector.Detect();
        diagnostics.Add(ToDiagnostic(report.Outlook));

        Step(15, "Collecting Outlook process evidence...");
        diagnostics.Add(CheckOutlookProcess(symptom));

        Step(25, "Discovering Outlook profiles...");
        report.Profiles = profileDiscovery.GetProfileInventory();
        diagnostics.Add(CheckProfiles(report.Outlook, report.Profiles));
        diagnostics.Add(CheckDefaultProfile(report.Profiles));

        Step(40, "Inspecting Outlook data files...");
        report.DataFiles = evidenceCollector.CollectDataFiles();
        diagnostics.Add(CheckDataFiles(report.Outlook, report.Profiles, report.DataFiles));

        Step(55, "Checking Windows time synchronization...");
        var timeResult = await powerShell.ExecuteAsync("w32tm /query /status", TimeSpan.FromSeconds(8), cancellationToken);
        diagnostics.Add(new TimeSyncParser().Parse(timeResult.Output + Environment.NewLine + timeResult.Error));

        Step(68, "Testing Microsoft 365 DNS and HTTPS endpoints...");
        diagnostics.AddRange(await CheckNetworkAsync(full, cancellationToken));

        if (full)
        {
            Step(78, "Reviewing Outlook event logs...");
            diagnostics.Add(await CheckEventLogsAsync(cancellationToken));
            Step(86, "Checking Windows service indicators...");
            diagnostics.Add(await CheckServicesAsync(cancellationToken));
        }

        Step(92, "Correlating findings and contradictions...");
        report.Contradictions = consistencyEngine.FindContradictions(report);
        foreach (var contradiction in report.Contradictions)
        {
            foreach (var diagnostic in diagnostics.Where(d => d.Status == DiagnosticStatus.Critical && contradiction.Title.Contains("Profile", StringComparison.OrdinalIgnoreCase)))
            {
                diagnostic.Status = DiagnosticStatus.Unknown;
                diagnostic.Severity = DiagnosticSeverity.Unknown;
                diagnostic.Finding = "Profile diagnostics are inconclusive because other evidence conflicts with a missing-profile conclusion.";
                diagnostic.Recommendation = "Validate Control Panel > Mail profile inventory before recreating the profile.";
            }
        }
        report.RootCauses = rootCauseEngine.Analyze(report);
        report.Health = scoringService.Score(diagnostics, report.Contradictions, report.RootCauses);
        report.EngineerSummary = summaryGenerator.Generate(report);

        Step(97, "Writing HTML and JSON reports...");
        reportGenerator.Write(report);
        sw.Stop();
        logger.Info($"Diagnostic completed in {sw.Elapsed}. Health={report.Health.HealthScore}. Html={report.HtmlReportPath}");
        Step(100, "Complete");
        return report;
    }

    private void Step(int percent, string text)
    {
        logger.Info(text);
        ProgressChanged?.Invoke(percent, text);
    }

    private static DiagnosticResult ToDiagnostic(OutlookClientEnvironment env)
    {
        var result = new DiagnosticResult
        {
            Id = "outlook.client.environment",
            Name = "Outlook Client Environment",
            Category = "Outlook",
            Status = env.ActiveClient == OutlookClientKind.None ? DiagnosticStatus.Warning : DiagnosticStatus.Info,
            Severity = env.ActiveClient == OutlookClientKind.None ? DiagnosticSeverity.Warning : DiagnosticSeverity.Info,
            Finding = $"Active client: {env.ActiveClient}. Classic installed: {env.ClassicInstalled}; running: {env.ClassicRunning}. New Outlook installed: {env.NewOutlookInstalled}; running: {env.NewOutlookRunning}.",
            Interpretation = "New Outlook installed is not treated as New Outlook active unless running evidence is present."
        };
        result.Evidence.AddRange(env.Evidence);
        return result;
    }

    private DiagnosticResult CheckOutlookProcess(UserSymptom symptom)
    {
        var procs = processInspector.GetProcesses("OUTLOOK");
        if (procs.Count == 0)
        {
            var status = symptom == UserSymptom.OutlookWontOpen ? DiagnosticStatus.Warning : DiagnosticStatus.Info;
            return new DiagnosticResult { Id = "outlook.process", Name = "Outlook Process", Category = "Outlook", Status = status, Severity = ToSeverity(status), Finding = "OUTLOOK.EXE is not currently running.", Interpretation = "Not running is informational unless the selected symptom is Outlook won't open.", Recommendation = status == DiagnosticStatus.Warning ? "Ask the user to start Outlook and capture the exact startup behavior or error." : string.Empty };
        }
        var hung = procs.Any(p => p.Responding == false);
        return new DiagnosticResult
        {
            Id = "outlook.process",
            Name = "Outlook Process",
            Category = "Outlook",
            Status = hung ? DiagnosticStatus.Critical : DiagnosticStatus.Pass,
            Severity = hung ? DiagnosticSeverity.Critical : DiagnosticSeverity.Pass,
            Finding = hung ? "OUTLOOK.EXE is running but not responding." : "OUTLOOK.EXE is running and responding.",
            Evidence = { new EvidenceItem("Process count", procs.Count.ToString(), hung ? "At least one process is hung." : "Classic Outlook process evidence is healthy.") },
            Recommendation = hung ? "Collect crash/hang event details and test Outlook safe mode before changing profile data." : string.Empty
        };
    }

    private static DiagnosticResult CheckProfiles(OutlookClientEnvironment env, ProfileInventory inventory)
    {
        if (env.ActiveClient == OutlookClientKind.NewOutlook)
            return new DiagnosticResult { Id = "outlook.profiles", Name = "Outlook Profile Discovery", Category = "Outlook Profile", Status = DiagnosticStatus.NotApplicable, Severity = DiagnosticSeverity.NotApplicable, Finding = "Classic Outlook profile registry diagnostics do not apply to New Outlook as the active client.", Interpretation = inventory.EvidenceSummary };
        if (inventory.Profiles.Count > 0)
            return new DiagnosticResult { Id = "outlook.profiles", Name = "Outlook Profile Discovery", Category = "Outlook Profile", Status = DiagnosticStatus.Pass, Severity = DiagnosticSeverity.Pass, Finding = $"{inventory.Profiles.Count} Outlook profile(s) discovered.", Interpretation = inventory.EvidenceSummary };
        if (env.ActiveClient == OutlookClientKind.Classic && !env.ClassicRunning)
            return new DiagnosticResult { Id = "outlook.profiles", Name = "Outlook Profile Discovery", Category = "Outlook Profile", Status = DiagnosticStatus.Critical, Severity = DiagnosticSeverity.Critical, Finding = "No Classic Outlook profile was discovered and Classic Outlook appears to require one.", Recommendation = "Open Control Panel > Mail > Show Profiles and verify whether a profile exists before creating a new one.", Interpretation = inventory.EvidenceSummary };
        return new DiagnosticResult { Id = "outlook.profiles", Name = "Outlook Profile Discovery", Category = "Outlook Profile", Status = DiagnosticStatus.Unknown, Severity = DiagnosticSeverity.Unknown, Finding = "Profile discovery was inconclusive.", Recommendation = "Verify Control Panel > Mail > Show Profiles; do not infer a missing profile from one registry path.", Interpretation = inventory.EvidenceSummary };
    }

    private static DiagnosticResult CheckDefaultProfile(ProfileInventory inventory)
    {
        if (!inventory.DefaultProfileConfigured)
            return new DiagnosticResult { Id = "outlook.profile.default", Name = "Default Profile Validation", Category = "Outlook Profile", Status = DiagnosticStatus.Unknown, Severity = DiagnosticSeverity.Unknown, Finding = "No configured default profile was found or it could not be determined.", Recommendation = "If the user is prompted to choose a profile, verify the default in Control Panel > Mail." };
        if (inventory.DefaultProfileExists == true)
            return new DiagnosticResult { Id = "outlook.profile.default", Name = "Default Profile Validation", Category = "Outlook Profile", Status = DiagnosticStatus.Pass, Severity = DiagnosticSeverity.Pass, Finding = $"Configured default profile '{inventory.ConfiguredDefaultProfile}' exists in the profile inventory." };
        return new DiagnosticResult { Id = "outlook.profile.default", Name = "Default Profile Validation", Category = "Outlook Profile", Status = DiagnosticStatus.Warning, Severity = DiagnosticSeverity.Warning, Finding = $"Configured default profile '{inventory.ConfiguredDefaultProfile}' was not found in discovered profiles.", Recommendation = "Validate whether this is a stale DefaultProfile value before changing profile configuration." };
    }

    private static DiagnosticResult CheckDataFiles(OutlookClientEnvironment env, ProfileInventory profiles, List<DataFileEvidence> files)
    {
        if (env.ActiveClient == OutlookClientKind.NewOutlook)
            return new DiagnosticResult { Id = "outlook.datafiles", Name = "OST/PST Data Files", Category = "OST/PST", Status = DiagnosticStatus.NotApplicable, Severity = DiagnosticSeverity.NotApplicable, Finding = "Classic OST/PST checks do not apply to New Outlook as the active client." };
        if (files.Count == 0)
        {
            var status = profiles.Profiles.Count > 0 && env.ActiveClient == OutlookClientKind.Classic ? DiagnosticStatus.Unknown : DiagnosticStatus.Info;
            return new DiagnosticResult { Id = "outlook.datafiles", Name = "OST/PST Data Files", Category = "OST/PST", Status = status, Severity = ToSeverity(status), Finding = "No OST/PST/NST files were found in the default local Outlook data folder.", Interpretation = "Absence is not automatically a failure; applicability depends on profile, account type, and cached mode." };
        }
        var stale = files.Where(f => !f.RecentlyModified).ToList();
        return new DiagnosticResult { Id = "outlook.datafiles", Name = "OST/PST Data Files", Category = "OST/PST", Status = stale.Count == files.Count ? DiagnosticStatus.Warning : DiagnosticStatus.Pass, Severity = stale.Count == files.Count ? DiagnosticSeverity.Warning : DiagnosticSeverity.Pass, Finding = $"{files.Count} Outlook data file(s) found; {files.Count - stale.Count} recently modified.", Recommendation = stale.Count == files.Count ? "Confirm Outlook is actively syncing before assuming the data file is unhealthy." : string.Empty };
    }

    private async Task<List<DiagnosticResult>> CheckNetworkAsync(bool full, CancellationToken token)
    {
        var endpoints = full ? new[] { "outlook.office365.com", "outlook.office.com", "login.microsoftonline.com", "autodiscover-s.outlook.com", "graph.microsoft.com" } : ["outlook.office365.com", "login.microsoftonline.com"];
        var results = new List<DiagnosticResult>();
        foreach (var endpoint in endpoints)
        {
            var dns = await powerShell.ExecuteAsync($"Resolve-DnsName {endpoint} -ErrorAction Stop | Select-Object -First 1 -ExpandProperty IPAddress", TimeSpan.FromSeconds(8), token);
            var dnsOk = dns.ExitCode == 0 && !string.IsNullOrWhiteSpace(dns.Output);
            results.Add(new DiagnosticResult { Id = $"network.dns.{endpoint}", Name = $"DNS Resolution - {endpoint}", Category = "Network", Status = dnsOk ? DiagnosticStatus.Pass : DiagnosticStatus.Warning, Severity = dnsOk ? DiagnosticSeverity.Pass : DiagnosticSeverity.Warning, Finding = dnsOk ? "DNS resolution succeeded." : "DNS resolution failed or returned no address.", RawResult = dns.Output + dns.Error, Recommendation = dnsOk ? string.Empty : "Validate DNS resolution for Microsoft 365 endpoints from the affected device." });
            if (dnsOk)
            {
                var tcp = await powerShell.ExecuteAsync($"Test-NetConnection {endpoint} -Port 443 -InformationLevel Quiet", TimeSpan.FromSeconds(10), token);
                var tcpOk = tcp.Output.Contains("True", StringComparison.OrdinalIgnoreCase);
                results.Add(new DiagnosticResult { Id = $"network.tcp443.{endpoint}", Name = $"TCP 443 - {endpoint}", Category = "Network", Status = tcpOk ? DiagnosticStatus.Pass : DiagnosticStatus.Warning, Severity = tcpOk ? DiagnosticSeverity.Pass : DiagnosticSeverity.Warning, Finding = tcpOk ? "TCP 443 connectivity succeeded." : "TCP 443 connectivity failed.", RawResult = tcp.Output + tcp.Error, Recommendation = tcpOk ? string.Empty : "Check firewall, proxy, VPN, or endpoint filtering for Microsoft 365 HTTPS traffic." });
            }
        }
        var proxy = await powerShell.ExecuteAsync("netsh winhttp show proxy", TimeSpan.FromSeconds(5), token);
        results.Add(new DiagnosticResult { Id = "network.proxy", Name = "WinHTTP Proxy", Category = "Network", Status = DiagnosticStatus.Info, Severity = DiagnosticSeverity.Info, Finding = proxy.Output.Contains("Direct access", StringComparison.OrdinalIgnoreCase) ? "No WinHTTP proxy is configured." : "WinHTTP proxy configuration detected.", RawResult = proxy.Output + proxy.Error, Recommendation = proxy.Output.Contains("Direct access", StringComparison.OrdinalIgnoreCase) ? string.Empty : "Confirm proxy configuration is expected and permits Microsoft 365 endpoints." });
        return results;
    }

    private async Task<DiagnosticResult> CheckEventLogsAsync(CancellationToken token)
    {
        var command = "Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-7)} -ErrorAction SilentlyContinue | Where-Object {$_.ProviderName -match 'Application Error|Application Hang|Windows Error Reporting' -and $_.Message -match 'OUTLOOK.EXE'} | Select-Object -First 20 TimeCreated,ProviderName,Id,Message | ConvertTo-Json -Compress";
        var result = await powerShell.ExecuteAsync(command, TimeSpan.FromSeconds(20), token);
        var found = !string.IsNullOrWhiteSpace(result.Output) && result.Output.Trim() != "null";
        return new DiagnosticResult { Id = "events.outlook.crashhang", Name = "Outlook Crash/Hang Events", Category = "Event Logs", Status = found ? DiagnosticStatus.Warning : DiagnosticStatus.Pass, Severity = found ? DiagnosticSeverity.Warning : DiagnosticSeverity.Pass, Finding = found ? "Recent Outlook crash/hang event evidence was found." : "No recent Outlook crash/hang events were found in Application log search.", RawResult = result.Output + result.Error, Recommendation = found ? "Review faulting module, exception code, and recurrence before disabling add-ins or rebuilding profile." : string.Empty };
    }

    private async Task<DiagnosticResult> CheckServicesAsync(CancellationToken token)
    {
        var result = await powerShell.ExecuteAsync("Get-Service ClickToRunSvc,VaultSvc,TokenBroker,CryptSvc,BITS -ErrorAction SilentlyContinue | Select Name,Status,StartType | ConvertTo-Json -Compress", TimeSpan.FromSeconds(8), token);
        return new DiagnosticResult { Id = "windows.services", Name = "Relevant Windows Services", Category = "System", Status = DiagnosticStatus.Info, Severity = DiagnosticSeverity.Info, Finding = "Relevant service states were collected. Manual/on-demand services are not treated as unhealthy only because they are stopped.", RawResult = result.Output + result.Error };
    }

    private static DiagnosticSeverity ToSeverity(DiagnosticStatus status) => Enum.Parse<DiagnosticSeverity>(status.ToString());
}


