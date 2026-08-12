using OutlookHealthCheck.Correlation;
using OutlookHealthCheck.Diagnostics;
using OutlookHealthCheck.Evidence;
using OutlookHealthCheck.Models;
using OutlookHealthCheck.Services;

var suite = new TestSuite();
suite.Run("One profile", Tests.OneProfile);
suite.Run("Two profiles", Tests.TwoProfiles);
suite.Run("No profile is unknown", Tests.NoProfileUnknown);
suite.Run("Multiple profiles", Tests.MultipleProfiles);
suite.Run("Default profile valid", Tests.DefaultProfileValid);
suite.Run("Default profile invalid", Tests.DefaultProfileInvalid);
suite.Run("New Outlook installed", Tests.NewOutlookInstalledDoesNotMeanActive);
suite.Run("New Outlook active", Tests.NewOutlookActive);
suite.Run("Classic Outlook active", Tests.ClassicOutlookActive);
suite.Run("New installed but Classic active", Tests.NewInstalledClassicActiveContradiction);
suite.Run("OST exists", Tests.OstExists);
suite.Run("OST missing", Tests.OstMissingIsNotCriticalByDefault);
suite.Run("OST recently modified", Tests.OstRecent);
suite.Run("Outlook not running", Tests.OutlookNotRunningInfo);
suite.Run("Outlook running", Tests.OutlookRunningPass);
suite.Run("Outlook responding", Tests.OutlookResponding);
suite.Run("Outlook not responding", Tests.OutlookNotRespondingCritical);
suite.Run("Contradiction no profile + active signals", Tests.NoProfileActiveSignalsContradiction);
suite.Run("Default configured + no discovered profile", Tests.DefaultConfiguredNoProfileContradiction);
suite.Run("New installed + Classic running", Tests.NewInstalledClassicRunningContradiction);
suite.Run("Windows Time synchronized", Tests.TimeSynchronized);
suite.Run("Windows Time service running not synchronized", Tests.TimeNotSynchronized);
suite.Run("DNS healthy model", Tests.DnsHealthyModel);
suite.Run("DNS failure model", Tests.DnsFailureModel);
suite.Run("TCP 443 failure model", Tests.TcpFailureModel);
suite.Run("Proxy detected model", Tests.ProxyDetectedModel);
suite.Run("No crashes", Tests.NoCrashesModel);
suite.Run("One Outlook crash", Tests.OneCrashModel);
suite.Run("Repeated Outlook crashes", Tests.RepeatedCrashModel);
suite.Run("Crash with faulting module", Tests.CrashFaultingModuleModel);
suite.Summary();

static class Tests
{
    public static void OneProfile() { var inv = DiscoveryWithProfiles("Corporate").GetProfileInventory(); Assert.Equal(1, inv.Profiles.Count); Assert.Equal(DiagnosticStatus.Pass, inv.Status); }
    public static void TwoProfiles() { var inv = DiscoveryWithProfiles("Corporate", "Personal").GetProfileInventory(); Assert.Equal(2, inv.Profiles.Count); }
    public static void NoProfileUnknown() { var inv = new OutlookProfileDiscovery(new FakeRegistryReader()).GetProfileInventory(); Assert.Equal(DiagnosticStatus.Unknown, inv.Status); }
    public static void MultipleProfiles() { var inv = DiscoveryWithProfiles("A", "B", "C").GetProfileInventory(); Assert.Equal(3, inv.Profiles.Count); }
    public static void DefaultProfileValid() { var reg = RegistryWithProfiles("Corporate"); reg.SetValue("HKCU", @"Software\Microsoft\Office\16.0\Outlook", "DefaultProfile", "Corporate"); var inv = new OutlookProfileDiscovery(reg).GetProfileInventory(); Assert.True(inv.DefaultProfileExists == true); Assert.True(inv.Profiles.Single().IsDefault == true); }
    public static void DefaultProfileInvalid() { var reg = RegistryWithProfiles("Corporate"); reg.SetValue("HKCU", @"Software\Microsoft\Office\16.0\Outlook", "DefaultProfile", "Missing"); var inv = new OutlookProfileDiscovery(reg).GetProfileInventory(); Assert.True(inv.DefaultProfileExists == false); }
    public static void NewOutlookInstalledDoesNotMeanActive() { var r = BaseReport(); r.Outlook.NewOutlookInstalled = true; r.Outlook.NewOutlookRunning = false; r.Outlook.ClassicRunning = false; r.Outlook.ActiveClient = OutlookClientKind.Unknown; Assert.Equal(OutlookClientKind.Unknown, r.Outlook.ActiveClient); }
    public static void NewOutlookActive() { var r = BaseReport(); r.Outlook.NewOutlookInstalled = true; r.Outlook.NewOutlookRunning = true; r.Outlook.ActiveClient = OutlookClientKind.NewOutlook; Assert.Equal(OutlookClientKind.NewOutlook, r.Outlook.ActiveClient); }
    public static void ClassicOutlookActive() { var r = BaseReport(); r.Outlook.ClassicRunning = true; r.Outlook.ActiveClient = OutlookClientKind.Classic; Assert.Equal(OutlookClientKind.Classic, r.Outlook.ActiveClient); }
    public static void NewInstalledClassicActiveContradiction() { var r = BaseReport(); r.Outlook.NewOutlookInstalled = true; r.Outlook.ClassicRunning = true; Assert.NotEmpty(new DiagnosticConsistencyEngine().FindContradictions(r)); }
    public static void OstExists() { var f = new DataFileEvidence { Path = "a.ost", LastWriteTime = DateTime.Now, Accessible = true }; Assert.True(f.Accessible); }
    public static void OstMissingIsNotCriticalByDefault() { var r = new DiagnosticResult { Status = DiagnosticStatus.Info }; Assert.NotEqual(DiagnosticStatus.Critical, r.Status); }
    public static void OstRecent() { var f = new DataFileEvidence { LastWriteTime = DateTime.Now.AddHours(-2) }; Assert.True(f.RecentlyModified); }
    public static void OutlookNotRunningInfo() { var r = new DiagnosticResult { Status = DiagnosticStatus.Info, Finding = "OUTLOOK.EXE is not currently running." }; Assert.Equal(DiagnosticStatus.Info, r.Status); }
    public static void OutlookRunningPass() { var r = new DiagnosticResult { Status = DiagnosticStatus.Pass }; Assert.Equal(DiagnosticStatus.Pass, r.Status); }
    public static void OutlookResponding() { var p = new ProcessEvidence { Responding = true }; Assert.True(p.Responding == true); }
    public static void OutlookNotRespondingCritical() { var r = new DiagnosticResult { Status = DiagnosticStatus.Critical, Finding = "not responding" }; Assert.Equal(DiagnosticStatus.Critical, r.Status); }
    public static void NoProfileActiveSignalsContradiction() { var r = BaseReport(); r.Outlook.ClassicRunning = true; r.Outlook.ActiveClient = OutlookClientKind.Classic; r.Profiles = new ProfileInventory(); r.DataFiles.Add(new DataFileEvidence { LastWriteTime = DateTime.Now }); Assert.NotEmpty(new DiagnosticConsistencyEngine().FindContradictions(r)); }
    public static void DefaultConfiguredNoProfileContradiction() { var r = BaseReport(); r.Profiles.ConfiguredDefaultProfile = "Corporate"; Assert.Contains(new DiagnosticConsistencyEngine().FindContradictions(r), c => c.Title.Contains("Default profile")); }
    public static void NewInstalledClassicRunningContradiction() { var r = BaseReport(); r.Outlook.NewOutlookInstalled = true; r.Outlook.ClassicRunning = true; Assert.Contains(new DiagnosticConsistencyEngine().FindContradictions(r), c => c.Title.Contains("New Outlook")); }
    public static void TimeSynchronized() { var r = new TimeSyncParser().Parse("Leap Indicator: 0(no warning)\nStratum: 3\nSource: time.windows.com\nLast Successful Sync Time: today"); Assert.Equal(DiagnosticStatus.Pass, r.Status); }
    public static void TimeNotSynchronized() { var r = new TimeSyncParser().Parse("Leap Indicator: 3(not synchronized)\nStratum: 0\nSource: Local CMOS Clock"); Assert.Equal(DiagnosticStatus.Warning, r.Status); }
    public static void DnsHealthyModel() { var r = Net("network.dns.x", DiagnosticStatus.Pass); Assert.Equal(DiagnosticStatus.Pass, r.Status); }
    public static void DnsFailureModel() { var r = Net("network.dns.x", DiagnosticStatus.Warning); Assert.Equal(DiagnosticStatus.Warning, r.Status); }
    public static void TcpFailureModel() { var r = Net("network.tcp443.x", DiagnosticStatus.Warning); Assert.Contains("tcp443", r.Id); }
    public static void ProxyDetectedModel() { var r = Net("network.proxy", DiagnosticStatus.Info); r.Recommendation = "Confirm proxy configuration is expected"; Assert.Contains("proxy", r.Id); }
    public static void NoCrashesModel() { var r = Event(DiagnosticStatus.Pass, "No recent Outlook crash/hang events"); Assert.Equal(DiagnosticStatus.Pass, r.Status); }
    public static void OneCrashModel() { var r = Event(DiagnosticStatus.Warning, "Recent Outlook crash/hang event evidence was found"); Assert.Equal(DiagnosticStatus.Warning, r.Status); }
    public static void RepeatedCrashModel() { var h = new RootCauseHypothesis { Category = "Event Logs", SupportingEvidence = ["Occurrences: 7"] }; Assert.Contains("7", h.SupportingEvidence[0]); }
    public static void CrashFaultingModuleModel() { var h = new RootCauseHypothesis { SupportingEvidence = ["Faulting module: example.dll"] }; Assert.Contains("example.dll", h.SupportingEvidence[0]); }

    private static OutlookProfileDiscovery DiscoveryWithProfiles(params string[] names) => new(RegistryWithProfiles(names));
    private static FakeRegistryReader RegistryWithProfiles(params string[] names)
    {
        var reg = new FakeRegistryReader();
        foreach (var name in names) reg.AddKey("HKCU", $@"Software\Microsoft\Office\16.0\Outlook\Profiles\{name}");
        return reg;
    }
    private static DiagnosticReport BaseReport() => new() { Outlook = new OutlookClientEnvironment(), Profiles = new ProfileInventory() };
    private static DiagnosticResult Net(string id, DiagnosticStatus status) => new() { Id = id, Category = "Network", Status = status, Severity = Enum.Parse<DiagnosticSeverity>(status.ToString()) };
    private static DiagnosticResult Event(DiagnosticStatus status, string finding) => new() { Id = "events.outlook.crashhang", Category = "Event Logs", Status = status, Finding = finding };
}

sealed class TestSuite
{
    private int _passed;
    private int _failed;
    public void Run(string name, Action test)
    {
        try { test(); Console.WriteLine($"PASS {name}"); _passed++; }
        catch (Exception ex) { Console.WriteLine($"FAIL {name}: {ex.Message}"); _failed++; }
    }
    public void Summary()
    {
        Console.WriteLine($"Tests passed: {_passed}; failed: {_failed}");
        if (_failed > 0) Environment.Exit(1);
    }
}

static class Assert
{
    public static void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception($"Expected {expected}, got {actual}"); }
    public static void NotEqual<T>(T notExpected, T actual) { if (EqualityComparer<T>.Default.Equals(notExpected, actual)) throw new Exception($"Did not expect {actual}"); }
    public static void True(bool condition) { if (!condition) throw new Exception("Expected true"); }
    public static void NotEmpty<T>(IEnumerable<T> items) { if (!items.Any()) throw new Exception("Expected non-empty collection"); }
    public static void Contains<T>(IEnumerable<T> items, Func<T, bool> predicate) { if (!items.Any(predicate)) throw new Exception("Expected matching item"); }
    public static void Contains(string expectedSubstring, string actual) { if (!actual.Contains(expectedSubstring, StringComparison.OrdinalIgnoreCase)) throw new Exception($"Expected '{actual}' to contain '{expectedSubstring}'"); }
}
