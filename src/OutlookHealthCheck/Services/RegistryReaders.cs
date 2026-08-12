using Microsoft.Win32;

namespace OutlookHealthCheck.Services;

public interface IRegistryReader
{
    bool KeyExists(string hive, string path);
    IReadOnlyList<string> GetSubKeyNames(string hive, string path);
    object? GetValue(string hive, string path, string name);
}

public sealed class WindowsRegistryReader : IRegistryReader
{
    public bool KeyExists(string hive, string path) => Open(hive, path) is not null;

    public IReadOnlyList<string> GetSubKeyNames(string hive, string path)
    {
        using var key = Open(hive, path);
        return key?.GetSubKeyNames() ?? [];
    }

    public object? GetValue(string hive, string path, string name)
    {
        using var key = Open(hive, path);
        return key?.GetValue(name);
    }

    private static RegistryKey? Open(string hive, string path)
    {
        try
        {
            var root = hive.Equals("HKCU", StringComparison.OrdinalIgnoreCase) ? Registry.CurrentUser :
                hive.Equals("HKLM", StringComparison.OrdinalIgnoreCase) ? Registry.LocalMachine : null;
            return root?.OpenSubKey(path, writable: false);
        }
        catch
        {
            return null;
        }
    }
}

public sealed class FakeRegistryReader : IRegistryReader
{
    private readonly Dictionary<string, Dictionary<string, object?>> _keys = new(StringComparer.OrdinalIgnoreCase);
    public void AddKey(string hive, string path) => _keys.TryAdd(Normalize(hive, path), new(StringComparer.OrdinalIgnoreCase));
    public void SetValue(string hive, string path, string name, object? value)
    {
        AddKey(hive, path);
        _keys[Normalize(hive, path)][name] = value;
    }
    public bool KeyExists(string hive, string path) => _keys.ContainsKey(Normalize(hive, path));
    public object? GetValue(string hive, string path, string name) => _keys.TryGetValue(Normalize(hive, path), out var values) && values.TryGetValue(name, out var value) ? value : null;
    public IReadOnlyList<string> GetSubKeyNames(string hive, string path)
    {
        var prefix = Normalize(hive, path).TrimEnd('\\') + "\\";
        return _keys.Keys
            .Where(k => k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Select(k => k[prefix.Length..])
            .Where(rest => rest.Length > 0 && !rest.Contains('\\'))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x)
            .ToList();
    }
    private static string Normalize(string hive, string path) => $"{hive}:\\{path.Trim('\\')}";
}
