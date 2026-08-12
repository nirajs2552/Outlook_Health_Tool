#requires -Version 5.1
<#
.SYNOPSIS
    CorrelationEngine.ps1 - Turns the flat list of check results into a
    health score, a set of root-cause hypotheses with confidence levels,
    and an engineer-friendly plain-English summary.

    Philosophy (spec section 51/32/33):
      - Never claim proof from correlation alone.
      - Classify each check into a root-cause category.
      - A root-cause hypothesis is "supported" when checks in its category
        show CRITICAL/WARNING/ERROR while adjacent categories (especially
        network/M365 service) are healthy - that pattern is what makes a
        LOCAL cause more likely than a SERVICE-side cause, and vice versa.
#>

# Map each Category used by the check modules to a root-cause bucket.
$Script:RootCauseCategoryMap = @{
    'Outlook'        = 'Local Outlook'
    'Outlook Profile'= 'Local Outlook'
    'OST/PST'        = 'Local Outlook'
    'Add-ins'        = 'Local Outlook'
    'Office'         = 'Office Installation'
    'System'         = 'Windows'
    'Windows Update'  = 'Windows'
    'Authentication' = 'Authentication'
    'Network'        = 'Network'
    'Event Logs'     = 'Windows'   # crash/hang evidence usually implicates local Outlook, handled specially below
}

function Get-HealthScore {
    param([Parameter(Mandatory)][object[]]$Results)

    $score = 100
    foreach ($r in $Results) {
        switch ($r.Severity) {
            'CRITICAL' { $score -= 20 }
            'ERROR'    { $score -= 15 }
            'WARNING'  { $score -= 7 }
            default    { } # PASS / INFO = no deduction
        }
    }
    if ($score -lt 0) { $score = 0 }
    if ($score -gt 100) { $score = 100 }
    return [int]$score
}

function Get-SeverityCounts {
    param([Parameter(Mandatory)][object[]]$Results)
    [PSCustomObject]@{
        Critical = @($Results | Where-Object { $_.Severity -eq 'CRITICAL' }).Count
        Warning  = @($Results | Where-Object { $_.Severity -eq 'WARNING' }).Count
        Passed   = @($Results | Where-Object { $_.Severity -eq 'PASS' }).Count
        Info     = @($Results | Where-Object { $_.Severity -eq 'INFO' }).Count
        Errors   = @($Results | Where-Object { $_.Severity -eq 'ERROR' }).Count
    }
}

function Get-RootCauseAnalysis {
    param([Parameter(Mandatory)][object[]]$Results)

    # Bucket every non-passing finding into a root-cause category.
    $buckets = @{}
    foreach ($r in $Results) {
        if ($r.Severity -notin @('WARNING','CRITICAL','ERROR')) { continue }
        $bucket = $Script:RootCauseCategoryMap[$r.Category]
        if (-not $bucket) { $bucket = 'Other' }

        # Special case: crash/hang evidence in Event Logs almost always implicates
        # local Outlook (profile/OST/add-in) rather than "Windows" broadly.
        if ($r.Category -eq 'Event Logs' -and $r.Check -match 'Crash|Hang') {
            $bucket = 'Local Outlook'
        }

        if (-not $buckets.ContainsKey($bucket)) { $buckets[$bucket] = New-Object System.Collections.Generic.List[object] }
        $buckets[$bucket].Add($r)
    }

    # Determine whether network/service-side is healthy - used to raise/lower
    # confidence for a "Local Outlook" hypothesis specifically.
    $networkIssues = @($Results | Where-Object { $_.Category -eq 'Network' -and $_.Severity -in @('WARNING','CRITICAL','ERROR') })
    $officeIssues  = @($Results | Where-Object { $_.Category -eq 'Office' -and $_.Severity -in @('WARNING','CRITICAL','ERROR') })
    $networkHealthy = $networkIssues.Count -eq 0
    $officeHealthy  = $officeIssues.Count -eq 0

    $hypotheses = @()
    foreach ($bucketName in $buckets.Keys) {
        $items = $buckets[$bucketName]
        $criticalCount = @($items | Where-Object { $_.Severity -eq 'CRITICAL' }).Count
        $warningCount  = @($items | Where-Object { $_.Severity -eq 'WARNING' }).Count

        # Base confidence purely from the severity mix in this bucket.
        $confidence = if ($criticalCount -ge 1) { 'MEDIUM' } elseif ($warningCount -ge 2) { 'LOW' } else { 'LOW' }

        # Raise confidence for Local Outlook / Office / Authentication when the
        # surrounding network and service layers are healthy - this is the
        # "process of elimination" logic requested in the spec.
        if ($bucketName -in @('Local Outlook','Office Installation','Authentication') -and $networkHealthy -and $criticalCount -ge 1) {
            $confidence = 'HIGH'
        } elseif ($bucketName -in @('Local Outlook','Office Installation','Authentication') -and $networkHealthy -and $warningCount -ge 2) {
            $confidence = 'MEDIUM'
        }

        if ($bucketName -eq 'Network' -and $criticalCount -ge 1) {
            $confidence = 'HIGH'
        }

        $evidence = $items | ForEach-Object { "$($_.Check): $($_.Finding)" }

        $hypotheses += [PSCustomObject]@{
            Category   = $bucketName
            Confidence = $confidence
            Evidence   = $evidence
            CriticalCount = $criticalCount
            WarningCount  = $warningCount
            Score      = ($criticalCount * 3) + $warningCount + ($(if ($confidence -eq 'HIGH') {5} elseif ($confidence -eq 'MEDIUM') {2} else {0}))
        }
    }

    # Rank by severity/confidence so the most likely cause comes first.
    $ranked = $hypotheses | Sort-Object -Property Score -Descending

    return [PSCustomObject]@{
        Hypotheses     = $ranked
        NetworkHealthy = $networkHealthy
        OfficeHealthy  = $officeHealthy
    }
}

function New-EngineerSummary {
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][object]$RootCause,
        [Parameter(Mandatory)][int]$HealthScore
    )

    $counts = Get-SeverityCounts -Results $Results
    $sb = New-Object System.Text.StringBuilder

    # Healthy-layer facts first (detected facts).
    $healthyFacts = @()
    if ($RootCause.NetworkHealthy) { $healthyFacts += 'Network and Microsoft 365 connectivity (DNS, HTTPS, Autodiscover) tested healthy.' }
    if ($RootCause.OfficeHealthy)  { $healthyFacts += 'Office installation and activation appear healthy.' }
    $outlookInstallOk = @($Results | Where-Object { $_.Check -eq 'Outlook Installation' -and $_.Severity -eq 'PASS' }).Count -gt 0
    if ($outlookInstallOk) { $healthyFacts += 'Outlook is installed and a valid executable was located.' }

    if ($healthyFacts.Count -gt 0) {
        [void]$sb.AppendLine(($healthyFacts -join ' '))
        [void]$sb.AppendLine('')
    }

    if ($counts.Critical -eq 0 -and $counts.Warning -eq 0) {
        [void]$sb.AppendLine('No critical issues or warnings were detected across system, Outlook, Office, identity, and network diagnostics. Outlook health appears normal based on the checks performed.')
    } else {
        $topHypothesis = $RootCause.Hypotheses | Select-Object -First 1
        if ($topHypothesis) {
            [void]$sb.AppendLine("The available evidence points primarily toward a $($topHypothesis.Category.ToLower()) issue (confidence: $($topHypothesis.Confidence)).")
        }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Detected: $($counts.Critical) critical finding(s) and $($counts.Warning) warning(s) out of $($Results.Count) checks performed.")
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Recommended next step:')
    $topRec = $Results | Where-Object { $_.Severity -eq 'CRITICAL' -and $_.Recommendation } | Select-Object -First 1
    if ($topRec) {
        [void]$sb.AppendLine($topRec.Recommendation)
    } elseif (($Results | Where-Object { $_.Severity -eq 'WARNING' -and $_.Recommendation } | Select-Object -First 1)) {
        [void]$sb.AppendLine(($Results | Where-Object { $_.Severity -eq 'WARNING' -and $_.Recommendation } | Select-Object -First 1).Recommendation)
    } else {
        [void]$sb.AppendLine('No further action indicated by automated diagnostics. Gather user-reported symptoms for targeted manual investigation.')
    }

    return $sb.ToString().Trim()
}

function Invoke-CorrelationAnalysis {
    param([Parameter(Mandatory)][object[]]$Results)

    $healthScore = Get-HealthScore -Results $Results
    $counts = Get-SeverityCounts -Results $Results
    $rootCause = Get-RootCauseAnalysis -Results $Results
    $summary = New-EngineerSummary -Results $Results -RootCause $rootCause -HealthScore $healthScore

    [PSCustomObject]@{
        HealthScore     = $healthScore
        Counts          = $counts
        RootCause       = $rootCause
        EngineerSummary = $summary
    }
}
