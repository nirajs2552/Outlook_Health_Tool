using System.Windows;
using OutlookHealthCheck.Correlation;
using OutlookHealthCheck.Diagnostics;
using OutlookHealthCheck.Evidence;
using OutlookHealthCheck.Models;
using OutlookHealthCheck.PowerShell;
using OutlookHealthCheck.Reporting;
using OutlookHealthCheck.Services;
using OutlookHealthCheck.UI;

namespace OutlookHealthCheck;

public partial class App : Application
{
    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        if (e.Args.Any(a => a.Equals("--quick", StringComparison.OrdinalIgnoreCase) || a.Equals("--full", StringComparison.OrdinalIgnoreCase)))
        {
            var full = e.Args.Any(a => a.Equals("--full", StringComparison.OrdinalIgnoreCase));
            var symptom = ParseSymptom(e.Args);
            var redactor = new SensitiveDataRedactor();
            var registry = new WindowsRegistryReader();
            var processes = new WindowsProcessInspector();
            var controller = new DiagnosticController(
                new OutlookEnvironmentDetector(registry, processes),
                new OutlookProfileDiscovery(registry),
                new LocalEvidenceCollector(new BasicPowerShellEvidence()),
                processes,
                new PowerShellExecutor(redactor),
                new DiagnosticConsistencyEngine(),
                new RootCauseEngine(),
                new HealthScoringService(),
                new EngineerSummaryGenerator(),
                new HtmlReportGenerator(redactor),
                new DiagnosticLogger(redactor));
            var report = await controller.RunAsync(symptom, full, CancellationToken.None);
            Console.WriteLine(report.HtmlReportPath);
            Shutdown(0);
            return;
        }

        new MainWindow().Show();
    }

    private static UserSymptom ParseSymptom(string[] args)
    {
        var raw = args.FirstOrDefault(a => a.StartsWith("--symptom=", StringComparison.OrdinalIgnoreCase))?.Split('=', 2)[1];
        return Enum.TryParse<UserSymptom>(raw, ignoreCase: true, out var symptom) ? symptom : UserSymptom.Unknown;
    }
}

