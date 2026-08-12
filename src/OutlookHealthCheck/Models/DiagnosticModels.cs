namespace OutlookHealthCheck.Models;

public enum DiagnosticStatus { Pass, Info, Warning, Critical, Error, Unknown, NotApplicable }
public enum DiagnosticSeverity { Pass, Info, Warning, Critical, Error, Unknown, NotApplicable }
public enum OutlookClientKind { Unknown, Classic, NewOutlook, Both, None }
public enum ConfidenceLevel { Unknown, Low, Medium, High }
public enum UserSymptom
{
    Unknown,
    OutlookWontOpen,
    OutlookCrashes,
    OutlookHangsFreezes,
    OutlookIsSlow,
    PasswordPrompts,
    CannotSendEmail,
    CannotReceiveEmail,
    OutlookDisconnected,
    SearchNotWorking,
    CalendarProblem,
    SharedMailboxProblem,
    AuthenticationSignInIssue,
    NewOutlookIssue,
    Other
}

public sealed record EvidenceItem(string Name, string RawEvidence, string Interpretation, string Source = "Local", bool IsSensitive = false);

public sealed class DiagnosticResult
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public DiagnosticStatus Status { get; set; } = DiagnosticStatus.Unknown;
    public DiagnosticSeverity Severity { get; set; } = DiagnosticSeverity.Unknown;
    public string Finding { get; set; } = string.Empty;
    public string Recommendation { get; set; } = string.Empty;
    public string Impact { get; set; } = string.Empty;
    public string Command { get; set; } = string.Empty;
    public string RawResult { get; set; } = string.Empty;
    public string Interpretation { get; set; } = string.Empty;
    public List<EvidenceItem> Evidence { get; init; } = [];
    public TimeSpan Duration { get; set; }
    public DateTimeOffset Timestamp { get; init; } = DateTimeOffset.Now;
}

public sealed class DiagnosticDefinition
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public IReadOnlySet<OutlookClientKind> ApplicableClients { get; init; } = new HashSet<OutlookClientKind>();
    public IReadOnlySet<UserSymptom> ApplicableSymptoms { get; init; } = new HashSet<UserSymptom>();
    public int SeverityWeight { get; init; }
    public IReadOnlyList<string> RequiredEvidence { get; init; } = [];
    public IReadOnlyList<string> Dependencies { get; init; } = [];
}

public sealed class OutlookClientEnvironment
{
    public bool ClassicInstalled { get; set; }
    public bool ClassicRunning { get; set; }
    public bool NewOutlookInstalled { get; set; }
    public bool NewOutlookRunning { get; set; }
    public string? ClassicExecutablePath { get; set; }
    public string? OfficeVersion { get; set; }
    public string? OfficeBuild { get; set; }
    public string? OfficeInstallationType { get; set; }
    public OutlookClientKind ActiveClient { get; set; } = OutlookClientKind.Unknown;
    public List<EvidenceItem> Evidence { get; } = [];
}

public sealed class EnvironmentEvidence
{
    public string ComputerName { get; set; } = Environment.MachineName;
    public string UserName { get; set; } = Environment.UserName;
    public string Domain { get; set; } = Environment.UserDomainName;
    public string OS { get; set; } = Environment.OSVersion.VersionString;
    public string Architecture { get; set; } = System.Runtime.InteropServices.RuntimeInformation.OSArchitecture.ToString();
    public string PowerShellVersion { get; set; } = string.Empty;
    public double? RamGb { get; set; }
    public string Cpu { get; set; } = string.Empty;
    public double? SystemDriveFreeGb { get; set; }
    public TimeSpan? Uptime { get; set; }
    public string TimeZone { get; set; } = TimeZoneInfo.Local.DisplayName;
}

public sealed class ProfileInventoryItem
{
    public string ProfileName { get; set; } = string.Empty;
    public string RegistryPath { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public bool? IsDefault { get; set; }
    public int? AccountCount { get; set; }
    public int? DataFileCount { get; set; }
    public string ProfileEvidence { get; set; } = string.Empty;
    public DiagnosticStatus Status { get; set; } = DiagnosticStatus.Unknown;
}

public sealed class ProfileInventory
{
    public List<ProfileInventoryItem> Profiles { get; init; } = [];
    public string? ConfiguredDefaultProfile { get; set; }
    public bool DefaultProfileConfigured => !string.IsNullOrWhiteSpace(ConfiguredDefaultProfile);
    public bool? DefaultProfileExists { get; set; }
    public DiagnosticStatus Status { get; set; } = DiagnosticStatus.Unknown;
    public string EvidenceSummary { get; set; } = string.Empty;
}

public sealed class DataFileEvidence
{
    public string Path { get; set; } = string.Empty;
    public string Extension { get; set; } = string.Empty;
    public double SizeGb { get; set; }
    public DateTime LastWriteTime { get; set; }
    public bool Accessible { get; set; }
    public bool RecentlyModified => DateTime.Now - LastWriteTime < TimeSpan.FromDays(3);
}

public sealed class Contradiction
{
    public string Title { get; init; } = string.Empty;
    public string Conclusion { get; init; } = string.Empty;
    public List<string> SupportingEvidence { get; init; } = [];
    public List<string> ContradictingEvidence { get; init; } = [];
}

public sealed class RootCauseHypothesis
{
    public string Category { get; init; } = string.Empty;
    public string Hypothesis { get; init; } = string.Empty;
    public ConfidenceLevel Confidence { get; init; } = ConfidenceLevel.Unknown;
    public List<string> SupportingEvidence { get; init; } = [];
    public List<string> ContradictingEvidence { get; init; } = [];
    public List<string> AffectedComponents { get; init; } = [];
    public string RecommendedNextStep { get; init; } = string.Empty;
}

public sealed class HealthScoreResult
{
    public int HealthScore { get; init; }
    public ConfidenceLevel DiagnosticConfidence { get; init; }
    public ConfidenceLevel RootCauseConfidence { get; init; }
    public Dictionary<string, int> Counts { get; init; } = [];
}

public sealed class DiagnosticReport
{
    public DateTimeOffset Timestamp { get; init; } = DateTimeOffset.Now;
    public UserSymptom Symptom { get; set; }
    public EnvironmentEvidence Environment { get; set; } = new();
    public OutlookClientEnvironment Outlook { get; set; } = new();
    public ProfileInventory Profiles { get; set; } = new();
    public List<DataFileEvidence> DataFiles { get; set; } = [];
    public List<DiagnosticResult> Diagnostics { get; set; } = [];
    public List<Contradiction> Contradictions { get; set; } = [];
    public List<RootCauseHypothesis> RootCauses { get; set; } = [];
    public HealthScoreResult Health { get; set; } = new();
    public string EngineerSummary { get; set; } = string.Empty;
    public string? HtmlReportPath { get; set; }
    public string? JsonReportPath { get; set; }
}

