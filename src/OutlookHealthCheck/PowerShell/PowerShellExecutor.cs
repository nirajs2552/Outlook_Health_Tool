using OutlookHealthCheck.Services;

namespace OutlookHealthCheck.PowerShell;

public sealed class PowerShellExecutionResult
{
    public string Command { get; init; } = string.Empty;
    public string Output { get; init; } = string.Empty;
    public string Error { get; init; } = string.Empty;
    public int? ExitCode { get; init; }
    public TimeSpan Duration { get; init; }
    public bool TimedOut { get; init; }
}

public interface IPowerShellExecutor
{
    Task<PowerShellExecutionResult> ExecuteAsync(string command, TimeSpan timeout, CancellationToken cancellationToken);
}

public sealed class PowerShellExecutor(IRedactor redactor) : IPowerShellExecutor
{
    public async Task<PowerShellExecutionResult> ExecuteAsync(string command, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(timeout);
        using var process = new System.Diagnostics.Process();
        process.StartInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -ExecutionPolicy Bypass -Command " + Quote(command),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        process.Start();
        var stdout = process.StandardOutput.ReadToEndAsync(timeoutCts.Token);
        var stderr = process.StandardError.ReadToEndAsync(timeoutCts.Token);
        try
        {
            await process.WaitForExitAsync(timeoutCts.Token).ConfigureAwait(false);
            sw.Stop();
            return new PowerShellExecutionResult
            {
                Command = command,
                Output = redactor.Redact(await stdout.ConfigureAwait(false)),
                Error = redactor.Redact(await stderr.ConfigureAwait(false)),
                ExitCode = process.ExitCode,
                Duration = sw.Elapsed,
                TimedOut = false
            };
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch { }
            sw.Stop();
            return new PowerShellExecutionResult { Command = command, Duration = sw.Elapsed, TimedOut = true, Error = "Diagnostic timed out." };
        }
    }

    private static string Quote(string command) => "\"" + command.Replace("\"", "\\\"") + "\"";
}
