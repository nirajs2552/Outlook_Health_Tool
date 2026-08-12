using OutlookHealthCheck.Models;
using OutlookHealthCheck.Services;

namespace OutlookHealthCheck.Evidence;

public sealed class OutlookEnvironmentDetector(IRegistryReader registry, IProcessInspector processes)
{
    public OutlookClientEnvironment Detect()
    {
        var env = new OutlookClientEnvironment();
        var classicPaths = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office", "root", "Office16", "OUTLOOK.EXE"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Office", "root", "Office16", "OUTLOOK.EXE"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office", "Office16", "OUTLOOK.EXE"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Office", "Office16", "OUTLOOK.EXE")
        };
        env.ClassicExecutablePath = classicPaths.FirstOrDefault(File.Exists);
        env.ClassicInstalled = env.ClassicExecutablePath is not null;
        env.Evidence.Add(new("Classic Outlook executable", env.ClassicExecutablePath ?? "Not found", env.ClassicInstalled ? "Classic Outlook is installed." : "Classic Outlook was not found in standard install paths."));

        var outlookProcesses = processes.GetProcesses("OUTLOOK");
        env.ClassicRunning = outlookProcesses.Any(p => (p.MainModulePath ?? string.Empty).Contains("Office", StringComparison.OrdinalIgnoreCase) || p.Name.Equals("OUTLOOK", StringComparison.OrdinalIgnoreCase));
        env.Evidence.Add(new("OUTLOOK.EXE process", $"Count={outlookProcesses.Count}", env.ClassicRunning ? "Classic Outlook is currently running." : "Classic Outlook is not currently running."));

        env.NewOutlookInstalled = registry.KeyExists("HKCU", @"Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages") ||
                                  registry.GetValue("HKCU", @"Software\Microsoft\Office\16.0\Outlook\Options\General", "NewOutlookInstalled") is not null;
        var newOutlookProcesses = processes.GetProcesses("olk").Concat(processes.GetProcesses("HxOutlook")).ToList();
        env.NewOutlookRunning = newOutlookProcesses.Count > 0;
        env.Evidence.Add(new("New Outlook indicators", $"InstalledIndicator={env.NewOutlookInstalled}; RunningProcessCount={newOutlookProcesses.Count}", env.NewOutlookRunning ? "New Outlook appears active." : "New Outlook package/flag alone is not treated as active."));

        var c2r = registry.GetValue("HKLM", @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");
        env.OfficeInstallationType = c2r is null ? "Unknown or MSI" : "Click-to-Run";
        env.OfficeVersion = File.Exists(env.ClassicExecutablePath) ? System.Diagnostics.FileVersionInfo.GetVersionInfo(env.ClassicExecutablePath!).ProductVersion : null;
        env.OfficeBuild = env.OfficeVersion;

        env.ActiveClient = (env.ClassicRunning, env.NewOutlookRunning) switch
        {
            (true, true) => OutlookClientKind.Both,
            (true, false) => OutlookClientKind.Classic,
            (false, true) => OutlookClientKind.NewOutlook,
            _ when env.ClassicInstalled || env.NewOutlookInstalled => OutlookClientKind.Unknown,
            _ => OutlookClientKind.None
        };
        return env;
    }
}

public sealed class OutlookProfileDiscovery(IRegistryReader registry)
{
    private static readonly string[] ProfileRoots =
    [
        @"Software\Microsoft\Office\16.0\Outlook\Profiles",
        @"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles",
        @"Software\Microsoft\Office\15.0\Outlook\Profiles"
    ];

    public ProfileInventory GetProfileInventory()
    {
        var inventory = new ProfileInventory();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var root in ProfileRoots)
        {
            var children = registry.GetSubKeyNames("HKCU", root);
            if (children.Count == 0)
            {
                inventory.EvidenceSummary += $"Registry location HKCU\\{root}: no profiles discovered. ";
                continue;
            }

            foreach (var profileName in children)
            {
                var key = profileName;
                if (!seen.Add(key)) continue;
                var accountCount = registry.GetSubKeyNames("HKCU", $@"{root}\{profileName}\9375CFF0413111d3B88A00104B2A6676").Count;
                var serviceCount = registry.GetSubKeyNames("HKCU", $@"{root}\{profileName}").Count;
                inventory.Profiles.Add(new ProfileInventoryItem
                {
                    ProfileName = profileName,
                    RegistryPath = $@"HKCU\{root}\{profileName}",
                    Source = root.Contains("Windows NT", StringComparison.OrdinalIgnoreCase) ? "Windows Messaging Subsystem" : "Office Outlook Profiles",
                    AccountCount = accountCount == 0 ? null : accountCount,
                    DataFileCount = serviceCount == 0 ? null : serviceCount,
                    ProfileEvidence = "Profile registry key exists.",
                    Status = DiagnosticStatus.Pass
                });
            }
        }

        inventory.ConfiguredDefaultProfile = registry.GetValue("HKCU", @"Software\Microsoft\Office\16.0\Outlook", "DefaultProfile")?.ToString();
        if (inventory.DefaultProfileConfigured)
        {
            inventory.DefaultProfileExists = inventory.Profiles.Any(p => p.ProfileName.Equals(inventory.ConfiguredDefaultProfile, StringComparison.OrdinalIgnoreCase));
            foreach (var p in inventory.Profiles) p.IsDefault = p.ProfileName.Equals(inventory.ConfiguredDefaultProfile, StringComparison.OrdinalIgnoreCase);
        }
        else
        {
            inventory.DefaultProfileExists = null;
            foreach (var p in inventory.Profiles) p.IsDefault = null;
        }

        inventory.Status = inventory.Profiles.Count > 0 ? DiagnosticStatus.Pass : DiagnosticStatus.Unknown;
        if (inventory.Profiles.Count == 0) inventory.EvidenceSummary += "No supported registry location produced a profile. This is inconclusive until active client and other evidence are considered.";
        return inventory;
    }
}

public sealed class LocalEvidenceCollector(IPowerShellEvidence powerShellEvidence)
{
    public EnvironmentEvidence CollectEnvironment()
    {
        var drive = DriveInfo.GetDrives().FirstOrDefault(d => d.IsReady && string.Equals(d.Name, Path.GetPathRoot(Environment.SystemDirectory), StringComparison.OrdinalIgnoreCase));
        return new EnvironmentEvidence
        {
            ComputerName = Environment.MachineName,
            UserName = Environment.UserName,
            Domain = Environment.UserDomainName,
            OS = Environment.OSVersion.VersionString,
            Architecture = System.Runtime.InteropServices.RuntimeInformation.OSArchitecture.ToString(),
            PowerShellVersion = powerShellEvidence.GetPowerShellVersion(),
            SystemDriveFreeGb = drive is null ? null : Math.Round(drive.AvailableFreeSpace / 1024d / 1024 / 1024, 1),
            Uptime = TimeSpan.FromMilliseconds(Environment.TickCount64),
            TimeZone = TimeZoneInfo.Local.DisplayName
        };
    }

    public List<DataFileEvidence> CollectDataFiles()
    {
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Outlook");
        if (!Directory.Exists(path)) return [];
        return Directory.EnumerateFiles(path, "*.*", SearchOption.AllDirectories)
            .Where(f => f.EndsWith(".ost", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".pst", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".nst", StringComparison.OrdinalIgnoreCase))
            .Select(f =>
            {
                var info = new FileInfo(f);
                return new DataFileEvidence { Path = f, Extension = info.Extension, SizeGb = Math.Round(info.Length / 1024d / 1024 / 1024, 2), LastWriteTime = info.LastWriteTime, Accessible = true };
            })
            .ToList();
    }
}

public interface IPowerShellEvidence { string GetPowerShellVersion(); }
public sealed class BasicPowerShellEvidence : IPowerShellEvidence { public string GetPowerShellVersion() => Environment.Version.ToString(); }
