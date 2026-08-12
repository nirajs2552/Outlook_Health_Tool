using System.Text.RegularExpressions;

namespace OutlookHealthCheck.Services;

public interface IRedactor { string Redact(string? text); }

public sealed class SensitiveDataRedactor : IRedactor
{
    private static readonly Regex[] Patterns =
    [
        new(@"(?i)(password|pwd|secret|token|refresh_token|access_token|apikey|api_key)\s*[:=]\s*\S+", RegexOptions.Compiled),
        new(@"\b[A-Za-z0-9\-_\.]{40,}\b", RegexOptions.Compiled),
        new(@"\b([A-Z0-9]{5}-){4}[A-Z0-9]{5}\b", RegexOptions.Compiled)
    ];

    public string Redact(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var result = text;
        foreach (var pattern in Patterns) result = pattern.Replace(result, "[REDACTED]");
        return result;
    }
}
