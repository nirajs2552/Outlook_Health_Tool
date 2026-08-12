using System.Diagnostics;
using OutlookHealthCheck.Models;

namespace OutlookHealthCheck.Services;

public interface IProcessInspector
{
    IReadOnlyList<ProcessEvidence> GetProcesses(string processName);
}

public sealed class ProcessEvidence
{
    public string Name { get; init; } = string.Empty;
    public int ProcessId { get; init; }
    public DateTime? StartTime { get; init; }
    public long WorkingSetBytes { get; init; }
    public bool? Responding { get; init; }
    public string? MainModulePath { get; init; }
    public string? CommandLine { get; init; }
}

public sealed class WindowsProcessInspector : IProcessInspector
{
    public IReadOnlyList<ProcessEvidence> GetProcesses(string processName)
    {
        var output = new List<ProcessEvidence>();
        foreach (var p in Process.GetProcessesByName(processName))
        {
            try
            {
                output.Add(new ProcessEvidence
                {
                    Name = p.ProcessName,
                    ProcessId = p.Id,
                    StartTime = Safe(() => p.StartTime),
                    WorkingSetBytes = Safe(() => p.WorkingSet64),
                    Responding = Safe(() => p.Responding),
                    MainModulePath = Safe(() => p.MainModule?.FileName),
                    CommandLine = null
                });
            }
            finally { p.Dispose(); }
        }
        return output;
    }

    private static T? Safe<T>(Func<T> read)
    {
        try { return read(); } catch { return default; }
    }
}

