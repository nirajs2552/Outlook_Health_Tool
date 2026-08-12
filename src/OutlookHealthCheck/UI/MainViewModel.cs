using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using OutlookHealthCheck.Correlation;
using OutlookHealthCheck.Diagnostics;
using OutlookHealthCheck.Evidence;
using OutlookHealthCheck.Models;
using OutlookHealthCheck.PowerShell;
using OutlookHealthCheck.Reporting;
using OutlookHealthCheck.Services;

namespace OutlookHealthCheck.UI;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly DiagnosticController _controller;
    private int _progress;
    private string _currentStatus = "Ready";
    private bool _isRunning;
    private DiagnosticReport? _report;
    private UserSymptom _selectedSymptom = UserSymptom.Unknown;
    private CancellationTokenSource? _cts;

    public MainViewModel()
    {
        var redactor = new SensitiveDataRedactor();
        var registry = new WindowsRegistryReader();
        var processes = new WindowsProcessInspector();
        var ps = new PowerShellExecutor(redactor);
        _controller = new DiagnosticController(
            new OutlookEnvironmentDetector(registry, processes),
            new OutlookProfileDiscovery(registry),
            new LocalEvidenceCollector(new BasicPowerShellEvidence()),
            processes,
            ps,
            new DiagnosticConsistencyEngine(),
            new RootCauseEngine(),
            new HealthScoringService(),
            new EngineerSummaryGenerator(),
            new HtmlReportGenerator(redactor),
            new DiagnosticLogger(redactor));
        _controller.ProgressChanged += (p, s) => App.Current.Dispatcher.Invoke(() => { Progress = p; CurrentStatus = s; });
        RunQuickCommand = new AsyncCommand(() => RunAsync(false), () => !IsRunning);
        RunFullCommand = new AsyncCommand(() => RunAsync(true), () => !IsRunning);
        CancelCommand = new RelayCommand(() => _cts?.Cancel(), () => IsRunning);
        OpenReportCommand = new RelayCommand(OpenReport, () => Report?.HtmlReportPath is not null);
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public Array Symptoms => Enum.GetValues<UserSymptom>();
    public UserSymptom SelectedSymptom { get => _selectedSymptom; set { _selectedSymptom = value; OnPropertyChanged(); } }
    public int Progress { get => _progress; set { _progress = value; OnPropertyChanged(); } }
    public string CurrentStatus { get => _currentStatus; set { _currentStatus = value; OnPropertyChanged(); } }
    public bool IsRunning { get => _isRunning; set { _isRunning = value; OnPropertyChanged(); CommandManager.InvalidateRequerySuggested(); } }
    public DiagnosticReport? Report { get => _report; set { _report = value; OnPropertyChanged(); RefreshCollections(); } }
    public ObservableCollection<DiagnosticResult> ImportantFindings { get; } = [];
    public ObservableCollection<DiagnosticResult> Diagnostics { get; } = [];
    public ObservableCollection<ProfileInventoryItem> Profiles { get; } = [];
    public ObservableCollection<Contradiction> Contradictions { get; } = [];
    public ICommand RunQuickCommand { get; }
    public ICommand RunFullCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand OpenReportCommand { get; }

    private async Task RunAsync(bool full)
    {
        IsRunning = true;
        Progress = 0;
        _cts = new CancellationTokenSource();
        try { Report = await _controller.RunAsync(SelectedSymptom, full, _cts.Token); }
        catch (OperationCanceledException) { CurrentStatus = "Cancelled"; }
        catch (Exception ex) { CurrentStatus = ex.Message; }
        finally { IsRunning = false; _cts.Dispose(); _cts = null; }
    }

    private void RefreshCollections()
    {
        ImportantFindings.Clear(); Diagnostics.Clear(); Profiles.Clear(); Contradictions.Clear();
        if (Report is null) return;
        foreach (var d in Report.Diagnostics.Where(d => d.Status is DiagnosticStatus.Critical or DiagnosticStatus.Warning or DiagnosticStatus.Unknown)) ImportantFindings.Add(d);
        foreach (var d in Report.Diagnostics) Diagnostics.Add(d);
        foreach (var p in Report.Profiles.Profiles) Profiles.Add(p);
        foreach (var c in Report.Contradictions) Contradictions.Add(c);
        OnPropertyChanged(nameof(Report));
        CommandManager.InvalidateRequerySuggested();
    }

    private void OpenReport()
    {
        if (Report?.HtmlReportPath is null) return;
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(Report.HtmlReportPath) { UseShellExecute = true });
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public sealed class RelayCommand(Action execute, Func<bool>? canExecute = null) : ICommand
{
    public event EventHandler? CanExecuteChanged { add => CommandManager.RequerySuggested += value; remove => CommandManager.RequerySuggested -= value; }
    public bool CanExecute(object? parameter) => canExecute?.Invoke() ?? true;
    public void Execute(object? parameter) => execute();
}

public sealed class AsyncCommand(Func<Task> execute, Func<bool>? canExecute = null) : ICommand
{
    public event EventHandler? CanExecuteChanged { add => CommandManager.RequerySuggested += value; remove => CommandManager.RequerySuggested -= value; }
    public bool CanExecute(object? parameter) => canExecute?.Invoke() ?? true;
    public async void Execute(object? parameter) => await execute();
}
