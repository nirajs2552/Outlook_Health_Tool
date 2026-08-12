#requires -Version 5.1
<#
.SYNOPSIS
    Outlook Health Check - Microsoft Outlook Diagnostic & Troubleshooting Assistant

.DESCRIPTION
    A read-only, diagnostic-first tool for IT/Messaging support engineers.
    Runs a comprehensive set of PowerShell-based checks against the local
    Outlook/Office/Windows configuration, correlates findings into a likely
    root cause, and produces a self-contained HTML report.

    This tool NEVER modifies the registry, deletes files, disables add-ins,
    removes credentials, or performs any other destructive action. It is
    strictly an assessment and reporting tool.

    Designed to run in the interactive signed-in user's own context (no
    elevation required, no Exchange Online admin connectivity attempted).

.USAGE
    .\OutlookHealthCheck.ps1

    If script execution is blocked by policy, run (per-session, does not
    change system-wide policy, and is the safest way to run an unsigned
    script you've reviewed):

        powershell.exe -ExecutionPolicy Bypass -File .\OutlookHealthCheck.ps1

    Or unblock the downloaded files first (Windows may mark files
    downloaded from the internet as blocked):

        Get-ChildItem -Recurse | Unblock-File
#>

[CmdletBinding()]
param()

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$Script:RootPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:ModulesPath = Join-Path $Script:RootPath 'Modules'
$Script:ReportsPath = Join-Path $Script:RootPath 'Reports'
$Script:LogsPath = Join-Path $Script:RootPath 'Logs'

# ----------------------------------------------------------------------------
# Load all diagnostic modules (dot-sourced, not .psm1, per project layout)
# ----------------------------------------------------------------------------
$moduleFiles = @(
    'Common.ps1',
    'SystemChecks.ps1',
    'OutlookChecks.ps1',
    'OfficeChecks.ps1',
    'IdentityChecks.ps1',
    'NetworkChecks.ps1',
    'EventLogChecks.ps1',
    'CorrelationEngine.ps1',
    'ReportGenerator.ps1'
)
foreach ($mf in $moduleFiles) {
    $path = Join-Path $Script:ModulesPath $mf
    if (Test-Path -LiteralPath $path) {
        . $path
    } else {
        [System.Windows.MessageBox]::Show("Required module not found: $path`n`nThe application cannot start.", 'Outlook Health Check - Fatal Error', 'OK', 'Error') | Out-Null
        exit 1
    }
}

Initialize-DiagLog -LogFolder $Script:LogsPath
if (-not (Test-Path -LiteralPath $Script:ReportsPath)) { New-Item -ItemType Directory -Path $Script:ReportsPath -Force | Out-Null }

# ----------------------------------------------------------------------------
# WPF GUI definition (XAML)
# ----------------------------------------------------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Outlook Health Check" Height="620" Width="760"
        WindowStartupLocation="CenterScreen" Background="#F4F6F8"
        FontFamily="Segoe UI">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#1E3A5F" Padding="28,20">
            <StackPanel>
                <TextBlock Text="Outlook Health Check" FontSize="24" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Microsoft Outlook Diagnostic &amp; Troubleshooting Assistant" FontSize="13" Foreground="#B8C9DC" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Body -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="28,20,28,10">
            <StackPanel>

                <Border Background="White" CornerRadius="8" Padding="18" BorderBrush="#E2E8F0" BorderThickness="1" Margin="0,0,0,16">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,10,0">
                            <TextBlock Text="USER" FontSize="10" Foreground="#64748B" FontWeight="Bold"/>
                            <TextBlock x:Name="txtUser" Text="-" FontSize="14" Margin="0,2,0,10"/>
                            <TextBlock Text="COMPUTER" FontSize="10" Foreground="#64748B" FontWeight="Bold"/>
                            <TextBlock x:Name="txtComputer" Text="-" FontSize="14" Margin="0,2,0,0"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="10,0,0,0">
                            <TextBlock Text="OUTLOOK" FontSize="10" Foreground="#64748B" FontWeight="Bold"/>
                            <TextBlock x:Name="txtOutlook" Text="-" FontSize="14" Margin="0,2,0,10"/>
                            <TextBlock Text="OPERATING SYSTEM" FontSize="10" Foreground="#64748B" FontWeight="Bold"/>
                            <TextBlock x:Name="txtOS" Text="-" FontSize="14" Margin="0,2,0,0"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                    <RadioButton x:Name="rbQuick" Content="Quick Check (~20s)" GroupName="mode" Margin="0,0,20,0" VerticalAlignment="Center"/>
                    <RadioButton x:Name="rbFull" Content="Full Check (comprehensive)" GroupName="mode" IsChecked="True" VerticalAlignment="Center"/>
                </StackPanel>

                <Button x:Name="btnRun" Content="RUN HEALTH CHECK" Height="52" FontSize="16" FontWeight="Bold"
                        Background="#1E3A5F" Foreground="White" BorderThickness="0" Cursor="Hand" Margin="0,0,0,18"/>

                <TextBlock Text="Progress" FontSize="12" Foreground="#64748B" Margin="0,0,0,4"/>
                <ProgressBar x:Name="progressBar" Height="22" Minimum="0" Maximum="100" Value="0" Margin="0,0,0,6"/>
                <TextBlock x:Name="txtCurrentCheck" Text="Idle - ready to run" FontSize="12" Foreground="#334155" Margin="0,0,0,20"/>

                <TextBlock x:Name="txtResultPath" Text="" FontSize="12" Foreground="#16a34a" TextWrapping="Wrap" Margin="0,0,0,10"/>

                <StackPanel Orientation="Horizontal">
                    <Button x:Name="btnOpenReport" Content="Open Report" Padding="14,8" Margin="0,0,10,0" IsEnabled="False"/>
                    <Button x:Name="btnOpenFolder" Content="Open Report Folder" Padding="14,8" Margin="0,0,10,0" IsEnabled="False"/>
                    <Button x:Name="btnCopySummary" Content="Copy Summary" Padding="14,8" Margin="0,0,10,0" IsEnabled="False"/>
                    <Button x:Name="btnExportJson" Content="Export JSON" Padding="14,8" IsEnabled="False"/>
                </StackPanel>

            </StackPanel>
        </ScrollViewer>

        <!-- Footer status bar -->
        <Border Grid.Row="2" Background="#1E3A5F" Padding="20,10">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Critical: " Foreground="#FCA5A5" FontWeight="Bold"/>
                <TextBlock x:Name="txtCritCount" Text="0" Foreground="White" Margin="0,0,20,0"/>
                <TextBlock Text="Warning: " Foreground="#FCD34D" FontWeight="Bold"/>
                <TextBlock x:Name="txtWarnCount" Text="0" Foreground="White" Margin="0,0,20,0"/>
                <TextBlock Text="Passed: " Foreground="#86EFAC" FontWeight="Bold"/>
                <TextBlock x:Name="txtPassCount" Text="0" Foreground="White" Margin="0,0,20,0"/>
                <TextBlock Text="Informational: " Foreground="#93C5FD" FontWeight="Bold"/>
                <TextBlock x:Name="txtInfoCount" Text="0" Foreground="White"/>
            </StackPanel>
        </Border>

    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Bind named elements
$controls = @{}
foreach ($name in @('txtUser','txtComputer','txtOutlook','txtOS','rbQuick','rbFull','btnRun',
                     'progressBar','txtCurrentCheck','txtResultPath','btnOpenReport','btnOpenFolder',
                     'btnCopySummary','btnExportJson','txtCritCount','txtWarnCount','txtPassCount','txtInfoCount')) {
    $controls[$name] = $window.FindName($name)
}

# ----------------------------------------------------------------------------
# Populate header info immediately on load
# ----------------------------------------------------------------------------
$sysSummaryResult = Get-SystemInfoSummary
$sysSummary = if ($sysSummaryResult.RawResult) { $null } else { $null }
try {
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $controls['txtOS'].Text = $osInfo.Caption
} catch { $controls['txtOS'].Text = 'Unknown' }

$controls['txtUser'].Text = "$env:USERDOMAIN\$env:USERNAME"
$controls['txtComputer'].Text = $env:COMPUTERNAME

$installInfo = Get-OutlookInstallationInfo
$controls['txtOutlook'].Text = if ($installInfo.Status -eq 'PASS') {
    ($installInfo.Value -split '\|')[1].Trim()
} else { 'Not detected' }

$Script:LastResults = $null
$Script:LastAnalysis = $null
$Script:LastReportPath = $null
$Script:LastJsonPath = $null

# ----------------------------------------------------------------------------
# Health check execution (runs on UI thread with DoEvents-style pumping via
# dispatcher so the progress bar updates smoothly without a background
# runspace - acceptable given the checks are short-lived, mostly local reads)
# ----------------------------------------------------------------------------
function Update-UiProgress {
    param([int]$Percent, [string]$CurrentCheck)
    $controls['progressBar'].Value = $Percent
    $controls['txtCurrentCheck'].Text = "Current Check: $CurrentCheck"
    $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Start-HealthCheck {
    param([bool]$QuickMode)

    $controls['btnRun'].IsEnabled = $false
    $controls['btnOpenReport'].IsEnabled = $false
    $controls['btnOpenFolder'].IsEnabled = $false
    $controls['btnCopySummary'].IsEnabled = $false
    $controls['btnExportJson'].IsEnabled = $false
    $controls['txtResultPath'].Text = ''

    $allResults = New-Object System.Collections.Generic.List[object]

    # Each entry: display name + scriptblock. Quick mode uses a reduced subset.
    $fullSteps = @(
        @{ Name = 'System Diagnostics';      Block = { Invoke-AllSystemChecks } }
        @{ Name = 'Outlook Installation & Profile'; Block = { Invoke-AllOutlookChecks } }
        @{ Name = 'Office Installation & Activation'; Block = { Invoke-AllOfficeChecks } }
        @{ Name = 'Identity & Authentication'; Block = { Invoke-AllIdentityChecks } }
        @{ Name = 'Network & Connectivity';   Block = { Invoke-AllNetworkChecks } }
        @{ Name = 'Event Logs & Update History'; Block = { Invoke-AllEventLogChecks } }
    )

    $quickSteps = @(
        @{ Name = 'System Diagnostics (quick)'; Block = { @(Test-SystemDiskSpace; Test-SystemUserPermissions) } }
        @{ Name = 'Outlook Process & Profile';   Block = { @(Get-OutlookInstallationInfo; Test-OutlookProcess; Test-OutlookProfiles; Test-OutlookDefaultProfile) } }
        @{ Name = 'Office Activation';           Block = { @(Get-OfficeActivationInfo) } }
        @{ Name = 'Identity';                    Block = { @(Get-M365IdentityInfo) } }
        @{ Name = 'Network & DNS';               Block = { @(Test-DnsResolution; Test-M365Connectivity) } }
    )

    $steps = if ($QuickMode) { $quickSteps } else { $fullSteps }
    $totalSteps = $steps.Count
    $stepIndex = 0

    foreach ($step in $steps) {
        $stepIndex++
        $pct = [int](($stepIndex / $totalSteps) * 100)
        Update-UiProgress -Percent $pct -CurrentCheck $step.Name

        $stepResults = & $step.Block
        foreach ($r in $stepResults) {
            if ($r) {
                $allResults.Add($r)
                $c = switch ($r.Severity) {
                    'CRITICAL' { $controls['txtCritCount'] }
                    'WARNING'  { $controls['txtWarnCount'] }
                    'PASS'     { $controls['txtPassCount'] }
                    default    { $controls['txtInfoCount'] }
                }
                if ($c) { $c.Text = [int]$c.Text + 1 }
            }
        }
        $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }

    Update-UiProgress -Percent 100 -CurrentCheck 'Analyzing results and generating report...'

    $analysis = Invoke-CorrelationAnalysis -Results $allResults.ToArray()

    $envInfo = @{
        ComputerName   = $env:COMPUTERNAME
        UserName       = "$env:USERDOMAIN\$env:USERNAME"
        OutlookVersion = $controls['txtOutlook'].Text
        OSCaption      = $controls['txtOS'].Text
    }

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $htmlPath = Join-Path $Script:ReportsPath "Outlook-Health-Report-$stamp.html"
    $jsonPath = Join-Path $Script:ReportsPath "Outlook-Health-Report-$stamp.json"

    New-DiagnosticHtmlReport -Results $allResults.ToArray() -Analysis $analysis -Environment $envInfo -OutputPath $htmlPath | Out-Null
    New-DiagnosticJsonReport -Results $allResults.ToArray() -Analysis $analysis -Environment $envInfo -OutputPath $jsonPath | Out-Null

    Write-DiagLog -Message "Report generated: $htmlPath (Health Score: $($analysis.HealthScore))" -Level 'INFO'

    $Script:LastResults = $allResults.ToArray()
    $Script:LastAnalysis = $analysis
    $Script:LastReportPath = $htmlPath
    $Script:LastJsonPath = $jsonPath

    $controls['txtCurrentCheck'].Text = "Complete - Health Score: $($analysis.HealthScore)/100"
    $controls['txtResultPath'].Text = "Report saved: $htmlPath"

    $controls['btnRun'].IsEnabled = $true
    $controls['btnOpenReport'].IsEnabled = $true
    $controls['btnOpenFolder'].IsEnabled = $true
    $controls['btnCopySummary'].IsEnabled = $true
    $controls['btnExportJson'].IsEnabled = $true

    # Auto-open the report in the default browser, per spec.
    try { Start-Process -FilePath $htmlPath } catch { Write-DiagLog -Message "Failed to auto-open report: $($_.Exception.Message)" -Level 'ERROR' }
}

# ----------------------------------------------------------------------------
# Event wiring
# ----------------------------------------------------------------------------
$controls['btnRun'].Add_Click({
    $controls['txtCritCount'].Text = '0'
    $controls['txtWarnCount'].Text = '0'
    $controls['txtPassCount'].Text = '0'
    $controls['txtInfoCount'].Text = '0'
    $quick = $controls['rbQuick'].IsChecked
    Start-HealthCheck -QuickMode $quick
})

$controls['btnOpenReport'].Add_Click({
    if ($Script:LastReportPath -and (Test-Path -LiteralPath $Script:LastReportPath)) {
        Start-Process -FilePath $Script:LastReportPath
    }
})

$controls['btnOpenFolder'].Add_Click({
    if (Test-Path -LiteralPath $Script:ReportsPath) {
        Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$Script:ReportsPath`""
    }
})

$controls['btnCopySummary'].Add_Click({
    if ($Script:LastAnalysis) {
        try {
            Set-Clipboard -Value $Script:LastAnalysis.EngineerSummary
            $controls['txtResultPath'].Text = 'Engineer summary copied to clipboard.'
        } catch { }
    }
})

$controls['btnExportJson'].Add_Click({
    if ($Script:LastJsonPath -and (Test-Path -LiteralPath $Script:LastJsonPath)) {
        Start-Process -FilePath 'explorer.exe' -ArgumentList "/select,`"$Script:LastJsonPath`""
    }
})

Write-DiagLog -Message "GUI initialized. Ready." -Level 'INFO'
$window.ShowDialog() | Out-Null
Write-DiagLog -Message "===== Outlook Health Check session ended =====" -Level 'INFO'

