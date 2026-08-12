using OutlookHealthCheck.Services;

namespace OutlookHealthCheck.Services;

public sealed class DiagnosticLogger(IRedactor redactor)
{
    public string LogDirectory { get; } = Path.Combine(AppContext.BaseDirectory, "Logs");
    public string LogPath { get; private set; } = string.Empty;

    public void Start()
    {
        Directory.CreateDirectory(LogDirectory);
        LogPath = Path.Combine(LogDirectory, $"OutlookHealthCheck-{DateTime.Now:yyyyMMdd-HHmmss}.log");
        Info("Diagnostic session started.");
    }

    public void Info(string message) => Write("INFO", message);
    public void Warn(string message) => Write("WARN", message);
    public void Error(string message) => Write("ERROR", message);

    private void Write(string level, string message)
    {
        if (string.IsNullOrWhiteSpace(LogPath)) Start();
        File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {redactor.Redact(message)}{Environment.NewLine}");
    }
}
