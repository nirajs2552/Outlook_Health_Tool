#requires -Version 5.1
<#
.SYNOPSIS
    EventLogChecks.ps1 - Windows Event Log analysis for Outlook/Office
    crashes and hangs (last ~7 days), plus Windows/Office update recency
    for timeline correlation. Read-only; summarizes rather than dumping
    raw events.
#>

function Get-OutlookCrashEvents {
    Invoke-SafeCheck -Category 'Event Logs' -Check 'Outlook Crash/Hang Detection' -ScriptBlock {
        $cmd = "Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1002;StartTime=(Get-Date).AddDays(-7)}"
        $startTime = (Get-Date).AddDays(-7)

        $events = @()
        try {
            $events = Get-WinEvent -FilterHashtable @{
                LogName   = 'Application'
                Id        = 1000, 1002   # 1000 = App Error (crash), 1002 = App Hang
                StartTime = $startTime
            } -ErrorAction Stop | Where-Object { $_.Message -match 'OUTLOOK\.EXE' }
        } catch [Exception] {
            if ($_.Exception.Message -notmatch 'No events were found') {
                return New-CheckResult -Category 'Event Logs' -Check 'Outlook Crash/Hang Detection' -Status 'ERROR' `
                    -Value 'Unable to query Application event log' -Finding $_.Exception.Message `
                    -Recommendation 'Verify the account running this tool has permission to read the Application event log.' -Command $cmd
            }
        }

        if (-not $events -or $events.Count -eq 0) {
            return New-CheckResult -Category 'Event Logs' -Check 'Outlook Crash/Hang Detection' -Status 'PASS' `
                -Value 'No OUTLOOK.EXE crash or hang events in the last 7 days' `
                -Expected 'No crash/hang events for Outlook' `
                -Finding 'No Outlook application errors or hangs found in the Application log for the last 7 days.' `
                -Command $cmd
        }

        $crashes = $events | Where-Object { $_.Id -eq 1000 }
        $hangs = $events | Where-Object { $_.Id -eq 1002 }
        $last24h = $events | Where-Object { $_.TimeCreated -ge (Get-Date).AddHours(-24) }

        $mostRecent = $events | Sort-Object TimeCreated -Descending | Select-Object -First 1
        $faultingModule = if ($mostRecent.Message -match 'Faulting module name:\s*(\S+)') { $Matches[1] } else { 'unknown' }
        $exceptionCode = if ($mostRecent.Message -match 'Exception code:\s*(\S+)') { $Matches[1] } else { 'unknown' }

        $status = if ($last24h.Count -ge 3 -or $crashes.Count -ge 5) { 'CRITICAL' }
                  elseif ($events.Count -ge 1) { 'WARNING' }
                  else { 'PASS' }

        $rec = switch ($status) {
            'CRITICAL' { "Repeated Outlook crashes/hangs detected ($($events.Count) events in 7 days, $($last24h.Count) in the last 24h). Investigate the faulting module ('$faultingModule') - this often points to a specific add-in or a corrupt profile/OST. Consider testing in Safe Mode and with a new profile." }
            'WARNING'  { "Outlook crash/hang event(s) detected in the last 7 days. Monitor for recurrence; if it continues, correlate with recently installed/updated add-ins." }
            default    { '' }
        }

        $summary = "$($crashes.Count) crash(es), $($hangs.Count) hang(s) in last 7 days. Most recent: $($mostRecent.TimeCreated) | Faulting module: $faultingModule | Exception: $exceptionCode"

        New-CheckResult -Category 'Event Logs' -Check 'Outlook Crash/Hang Detection' -Status $status `
            -Value $summary -Expected 'No repeated crash/hang pattern' `
            -Finding "$(if($status -eq 'PASS'){'No significant crash pattern.'}else{'A crash/hang pattern was detected for OUTLOOK.EXE.'})" `
            -Recommendation $rec -Command $cmd `
            -RawResult ($events | Select-Object TimeCreated, Id, @{n='Summary';e={($_.Message -split "`n")[0]}} | Sort-Object TimeCreated -Descending | Format-Table -AutoSize | Out-String)
    }
}

function Get-OfficeApplicationErrors {
    Invoke-SafeCheck -Category 'Event Logs' -Check 'Office Application Errors' -ScriptBlock {
        $cmd = "Get-WinEvent -FilterHashtable @{LogName='Application';ProviderName='Microsoft Office*';StartTime=(Get-Date).AddDays(-7)}"
        $startTime = (Get-Date).AddDays(-7)
        $events = @()
        try {
            $events = Get-WinEvent -FilterHashtable @{
                LogName   = 'Application'
                StartTime = $startTime
            } -ErrorAction Stop | Where-Object { $_.ProviderName -match 'Office|MsiInstaller' -and $_.LevelDisplayName -eq 'Error' }
        } catch { }

        if (-not $events -or $events.Count -eq 0) {
            return New-CheckResult -Category 'Event Logs' -Check 'Office Application Errors' -Status 'PASS' `
                -Value 'No Office-related error events in the last 7 days' -Expected 'No recurring Office error events' `
                -Finding 'No relevant Office error-level events found in the Application log.' -Command $cmd
        }

        $status = if ($events.Count -ge 10) { 'WARNING' } else { 'INFO' }
        $topProviders = $events | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 3
        $summary = "$($events.Count) Office-related error event(s) in last 7 days. Top sources: " + (($topProviders | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ')

        New-CheckResult -Category 'Event Logs' -Check 'Office Application Errors' -Status $status `
            -Value $summary -Expected 'Low volume of Office-related error events' `
            -Finding "$(if($status -eq 'INFO'){'Some Office-related error events found; volume is not alarming on its own.'}else{'A notable volume of Office-related error events was found.'})" `
            -Command $cmd -RawResult ($events | Select-Object TimeCreated, ProviderName, Id | Format-Table -AutoSize | Out-String)
    }
}

function Get-WindowsUpdateRecency {
    Invoke-SafeCheck -Category 'Windows Update' -Check 'Recent Windows Updates' -ScriptBlock {
        $cmd = 'Get-HotFix | Sort-Object InstalledOn -Descending'
        $hotfixes = Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending
        if (-not $hotfixes -or $hotfixes.Count -eq 0) {
            return New-CheckResult -Category 'Windows Update' -Check 'Recent Windows Updates' -Status 'INFO' `
                -Value 'No hotfix history available via Get-HotFix' -Expected 'Informational' `
                -Finding 'Could not enumerate Windows Update history via Get-HotFix on this system.' -Command $cmd
        }

        $mostRecent = $hotfixes | Select-Object -First 1
        $daysAgo = if ($mostRecent.InstalledOn) { [math]::Round(((Get-Date) - $mostRecent.InstalledOn).TotalDays, 1) } else { $null }

        $finding = if ($null -ne $daysAgo -and $daysAgo -le 3) {
            "A Windows update ($($mostRecent.HotFixID)) was installed approximately $daysAgo day(s) ago. This is a temporal correlation only and does not by itself establish causation - note it if the reported Outlook issue began around the same time."
        } else {
            'Most recent Windows update installation is not within the last 3 days.'
        }

        New-CheckResult -Category 'Windows Update' -Check 'Recent Windows Updates' -Status 'INFO' `
            -Value "Most recent: $($mostRecent.HotFixID) installed $($mostRecent.InstalledOn)" `
            -Expected 'Informational - used for timeline correlation, not a pass/fail check' `
            -Finding $finding -Command $cmd -RawResult ($hotfixes | Select-Object -First 5 | Format-Table -AutoSize | Out-String)
    }
}

function Invoke-AllEventLogChecks {
    @(
        Get-OutlookCrashEvents
        Get-OfficeApplicationErrors
        Get-WindowsUpdateRecency
    )
}
