#requires -Version 5.1
<#
.SYNOPSIS
    ReportGenerator.ps1 - Builds the self-contained HTML diagnostic report
    (all CSS/JS inline, no external dependencies, no internet required to
    view) and an equivalent JSON export for automation.
#>

function ConvertTo-HtmlSafe {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-SeverityBadgeClass {
    param([string]$Severity)
    switch ($Severity) {
        'PASS'     { 'badge-pass' }
        'INFO'     { 'badge-info' }
        'WARNING'  { 'badge-warning' }
        'CRITICAL' { 'badge-critical' }
        'ERROR'    { 'badge-error' }
        default    { 'badge-info' }
    }
}

function New-DiagnosticHtmlReport {
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][object]$Analysis,
        [Parameter(Mandatory)][hashtable]$Environment,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $counts = $Analysis.Counts
    $score = $Analysis.HealthScore
    $scoreColor = if ($score -ge 80) { '#16a34a' } elseif ($score -ge 50) { '#d97706' } else { '#dc2626' }

    $genTime = Get-Date -Format 'dddd, MMMM d, yyyy - HH:mm:ss'

    # ---- Critical findings section -----------------------------------
    $criticalItems = $Results | Where-Object { $_.Severity -eq 'CRITICAL' }
    $criticalHtml = if ($criticalItems.Count -eq 0) {
        '<div class="no-critical">✅ No critical findings detected.</div>'
    } else {
        ($criticalItems | ForEach-Object {
            @"
<div class="critical-card">
  <div class="critical-card-header">🔴 CRITICAL &mdash; $(ConvertTo-HtmlSafe $_.Check)</div>
  <div class="critical-card-body">
    <p class="finding-text">$(ConvertTo-HtmlSafe $_.Finding)</p>
    <table class="evidence-table">
      <tr><th>Expected</th><td>$(ConvertTo-HtmlSafe $_.Expected)</td></tr>
      <tr><th>Actual</th><td>$(ConvertTo-HtmlSafe $_.Value)</td></tr>
    </table>
    $(if ($_.Recommendation) { "<p class='recommendation'><strong>Recommendation:</strong> $(ConvertTo-HtmlSafe $_.Recommendation)</p>" })
  </div>
</div>
"@
        }) -join "`n"
    }

    # ---- Root cause section --------------------------------------------
    $rootCauseHtml = if ($Analysis.RootCause.Hypotheses.Count -eq 0) {
        '<p>No significant abnormalities were correlated across categories.</p>'
    } else {
        ($Analysis.RootCause.Hypotheses | Select-Object -First 3 | ForEach-Object {
            $confClass = "conf-$($_.Confidence.ToLower())"
            @"
<div class="hypothesis-card">
  <div class="hypothesis-header">
    <span class="hypothesis-title">$(ConvertTo-HtmlSafe $_.Category)</span>
    <span class="confidence-badge $confClass">$($_.Confidence) CONFIDENCE</span>
  </div>
  <ul class="evidence-list">
    $(($_.Evidence | ForEach-Object { "<li>$(ConvertTo-HtmlSafe $_)</li>" }) -join "`n")
  </ul>
</div>
"@
        }) -join "`n"
    }

    # ---- Category sections (collapsible) --------------------------------
    $categories = $Results | Group-Object Category | Sort-Object Name
    $categorySectionsHtml = ($categories | ForEach-Object {
        $catName = $_.Name
        $catResults = $_.Group
        $catCritical = @($catResults | Where-Object { $_.Severity -eq 'CRITICAL' }).Count
        $catWarning  = @($catResults | Where-Object { $_.Severity -eq 'WARNING' }).Count

        $rowsHtml = ($catResults | ForEach-Object {
            $badgeClass = Get-SeverityBadgeClass -Severity $_.Severity
            $rowClass = "row-$($_.Severity.ToLower())"
            @"
<div class="check-row $rowClass">
  <div class="check-row-summary" onclick="toggleDetails(this)">
    <span class="badge $badgeClass">$($_.Severity)</span>
    <span class="check-name">$(ConvertTo-HtmlSafe $_.Check)</span>
    <span class="check-value">$(ConvertTo-HtmlSafe $_.Value)</span>
    <span class="expand-arrow">▸</span>
  </div>
  <div class="check-row-details">
    <table class="details-table">
      <tr><th>Finding</th><td>$(ConvertTo-HtmlSafe $_.Finding)</td></tr>
      <tr><th>Expected</th><td>$(ConvertTo-HtmlSafe $_.Expected)</td></tr>
      $(if ($_.Recommendation) { "<tr><th>Recommendation</th><td>$(ConvertTo-HtmlSafe $_.Recommendation)</td></tr>" })
      <tr><th>PowerShell Command</th><td><code>$(ConvertTo-HtmlSafe $_.Command)</code></td></tr>
      $(if ($_.RawResult) { "<tr><th>Raw Result</th><td><pre>$(ConvertTo-HtmlSafe $_.RawResult)</pre></td></tr>" })
    </table>
  </div>
</div>
"@
        }) -join "`n"

        @"
<div class="category-section">
  <div class="category-header" onclick="toggleCategory(this)">
    <span class="category-arrow">▼</span>
    <span class="category-title">$(ConvertTo-HtmlSafe $catName)</span>
    <span class="category-counts">
      $(if ($catCritical -gt 0) { "<span class='mini-badge mini-critical'>$catCritical critical</span>" })
      $(if ($catWarning -gt 0) { "<span class='mini-badge mini-warning'>$catWarning warning</span>" })
      <span class="mini-badge mini-total">$($catResults.Count) checks</span>
    </span>
  </div>
  <div class="category-body">
    $rowsHtml
  </div>
</div>
"@
    }) -join "`n"

    # ---- Full HTML document ------------------------------------------
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Outlook Health Check Report - $(ConvertTo-HtmlSafe $Environment.ComputerName)</title>
<style>
  :root {
    --pass: #16a34a; --info: #2563eb; --warning: #d97706; --critical: #dc2626; --error: #991b1b;
    --bg: #f4f6f8; --card-bg: #ffffff; --border: #e2e8f0; --text: #1e293b; --muted: #64748b;
  }
  * { box-sizing: border-box; }
  body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: var(--bg); color: var(--text); margin: 0; padding: 0; line-height: 1.5;
  }
  .header {
    background: linear-gradient(135deg, #1e3a5f 0%, #0f2942 100%);
    color: white; padding: 28px 40px;
  }
  .header h1 { margin: 0; font-size: 26px; font-weight: 600; }
  .header .subtitle { color: #b8c9dc; font-size: 14px; margin-top: 4px; }
  .header-meta { display: flex; flex-wrap: wrap; gap: 24px; margin-top: 18px; font-size: 13px; }
  .header-meta div span.label { color: #93aec8; display: block; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px;}
  .header-meta div span.value { font-weight: 600; }

  .container { max-width: 1100px; margin: 0 auto; padding: 24px 40px 60px; }

  .summary-bar {
    display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap;
  }
  .score-card {
    background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px;
    padding: 20px 28px; flex: 1; min-width: 220px; text-align: center;
  }
  .score-number { font-size: 42px; font-weight: 700; color: $scoreColor; }
  .score-label { color: var(--muted); font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }
  .counts-card {
    background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px;
    padding: 20px 28px; flex: 2; min-width: 320px; display: flex; justify-content: space-around;
  }
  .count-item { text-align: center; }
  .count-number { font-size: 28px; font-weight: 700; }
  .count-label { font-size: 12px; color: var(--muted); text-transform: uppercase; }
  .count-critical .count-number { color: var(--critical); }
  .count-warning .count-number { color: var(--warning); }
  .count-passed .count-number { color: var(--pass); }
  .count-info .count-number { color: var(--info); }

  .section-title {
    font-size: 18px; font-weight: 600; margin: 32px 0 12px; color: #1e3a5f;
    border-bottom: 2px solid var(--border); padding-bottom: 8px;
  }

  .engineer-summary {
    background: #eef4fb; border-left: 4px solid #1e3a5f; border-radius: 6px;
    padding: 18px 22px; white-space: pre-wrap; font-size: 14px;
  }

  .no-critical {
    background: #ecfdf5; border: 1px solid #a7f3d0; color: #065f46;
    padding: 16px 20px; border-radius: 8px; font-weight: 600;
  }
  .critical-card {
    background: #fef2f2; border: 2px solid var(--critical); border-radius: 8px;
    margin-bottom: 14px; overflow: hidden;
  }
  .critical-card-header {
    background: var(--critical); color: white; font-weight: 700; padding: 10px 18px; font-size: 14px;
  }
  .critical-card-body { padding: 14px 18px; }
  .finding-text { font-weight: 600; margin: 0 0 10px; }
  .evidence-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 10px;}
  .evidence-table th { text-align: left; color: var(--muted); padding: 4px 10px 4px 0; width: 120px; vertical-align: top;}
  .evidence-table td { padding: 4px 0; }
  .recommendation { background: white; border-radius: 6px; padding: 10px 14px; font-size: 13px; margin: 0; }

  .hypothesis-card {
    background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px;
    padding: 16px 20px; margin-bottom: 12px;
  }
  .hypothesis-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
  .hypothesis-title { font-weight: 700; font-size: 15px; }
  .confidence-badge { font-size: 11px; font-weight: 700; padding: 3px 10px; border-radius: 20px; letter-spacing: 0.5px;}
  .conf-high { background: #fee2e2; color: var(--critical); }
  .conf-medium { background: #fef3c7; color: var(--warning); }
  .conf-low { background: #e0e7ff; color: #4338ca; }
  .evidence-list { margin: 6px 0 0; padding-left: 20px; font-size: 13px; color: #334155; }
  .evidence-list li { margin-bottom: 4px; }

  .category-section {
    background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px;
    margin-bottom: 10px; overflow: hidden;
  }
  .category-header {
    display: flex; align-items: center; gap: 10px; padding: 12px 18px; cursor: pointer;
    background: #f8fafc; user-select: none;
  }
  .category-header:hover { background: #f1f5f9; }
  .category-arrow { transition: transform 0.15s; font-size: 12px; color: var(--muted); }
  .category-header.collapsed .category-arrow { transform: rotate(-90deg); }
  .category-title { font-weight: 600; flex: 1; }
  .category-counts { display: flex; gap: 6px; }
  .mini-badge { font-size: 11px; padding: 2px 8px; border-radius: 10px; font-weight: 600; }
  .mini-critical { background: #fee2e2; color: var(--critical); }
  .mini-warning { background: #fef3c7; color: var(--warning); }
  .mini-total { background: #f1f5f9; color: var(--muted); }
  .category-body { padding: 4px 0; }
  .category-body.collapsed { display: none; }

  .check-row { border-top: 1px solid var(--border); }
  .check-row-summary {
    display: flex; align-items: center; gap: 12px; padding: 10px 18px; cursor: pointer;
  }
  .check-row-summary:hover { background: #fafbfc; }
  .badge {
    font-size: 11px; font-weight: 700; padding: 3px 9px; border-radius: 4px; min-width: 62px;
    text-align: center; color: white;
  }
  .badge-pass { background: var(--pass); }
  .badge-info { background: var(--info); }
  .badge-warning { background: var(--warning); }
  .badge-critical { background: var(--critical); }
  .badge-error { background: var(--error); }
  .check-name { font-weight: 600; min-width: 220px; }
  .check-value { color: var(--muted); font-size: 13px; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .expand-arrow { color: var(--muted); font-size: 12px; transition: transform 0.15s; }
  .check-row-summary.expanded .expand-arrow { transform: rotate(90deg); }
  .check-row-details { display: none; padding: 4px 18px 16px 84px; background: #fafbfc; }
  .check-row-details.visible { display: block; }
  .details-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  .details-table th { text-align: left; color: var(--muted); width: 160px; vertical-align: top; padding: 6px 10px 6px 0; }
  .details-table td { padding: 6px 0; }
  .details-table code { background: #eef2f7; padding: 2px 6px; border-radius: 4px; font-family: Consolas, monospace; font-size: 12px; }
  .details-table pre { background: #0f172a; color: #e2e8f0; padding: 10px 14px; border-radius: 6px; font-size: 12px; overflow-x: auto; white-space: pre-wrap; }

  .footer { text-align: center; color: var(--muted); font-size: 12px; margin-top: 40px; }
  .row-warning .badge { }
  .toolbar { margin: 10px 0 20px; display: flex; gap: 10px; }
  .toolbar button {
    background: #1e3a5f; color: white; border: none; padding: 8px 16px; border-radius: 6px;
    font-size: 13px; cursor: pointer;
  }
  .toolbar button:hover { background: #16324e; }
</style>
</head>
<body>

<div class="header">
  <h1>🩺 Outlook Health Check</h1>
  <div class="subtitle">Microsoft Outlook Diagnostic &amp; Troubleshooting Assistant &mdash; Diagnostic Report</div>
  <div class="header-meta">
    <div><span class="label">Computer</span><span class="value">$(ConvertTo-HtmlSafe $Environment.ComputerName)</span></div>
    <div><span class="label">User</span><span class="value">$(ConvertTo-HtmlSafe $Environment.UserName)</span></div>
    <div><span class="label">Outlook</span><span class="value">$(ConvertTo-HtmlSafe $Environment.OutlookVersion)</span></div>
    <div><span class="label">Windows</span><span class="value">$(ConvertTo-HtmlSafe $Environment.OSCaption)</span></div>
    <div><span class="label">Generated</span><span class="value">$genTime</span></div>
  </div>
</div>

<div class="container">

  <div class="summary-bar">
    <div class="score-card">
      <div class="score-number">$score</div>
      <div class="score-label">Overall Health Score / 100</div>
    </div>
    <div class="counts-card">
      <div class="count-item count-critical"><div class="count-number">$($counts.Critical)</div><div class="count-label">Critical</div></div>
      <div class="count-item count-warning"><div class="count-number">$($counts.Warning)</div><div class="count-label">Warning</div></div>
      <div class="count-item count-passed"><div class="count-number">$($counts.Passed)</div><div class="count-label">Passed</div></div>
      <div class="count-item count-info"><div class="count-number">$($counts.Info)</div><div class="count-label">Info</div></div>
    </div>
  </div>

  <div class="section-title">Engineer Summary</div>
  <div class="engineer-summary">$(ConvertTo-HtmlSafe $Analysis.EngineerSummary)</div>

  <div class="section-title">Critical Findings</div>
  $criticalHtml

  <div class="section-title">Likely Root Cause</div>
  $rootCauseHtml

  <div class="section-title">Diagnostic Details by Category</div>
  <div class="toolbar">
    <button onclick="expandAll()">Expand All</button>
    <button onclick="collapseAll()">Collapse All</button>
  </div>
  $categorySectionsHtml

  <div class="footer">
    Generated by Outlook Health Check &middot; Read-only diagnostic tool &middot; No configuration was modified during this scan.
  </div>

</div>

<script>
function toggleDetails(el) {
  el.classList.toggle('expanded');
  var details = el.nextElementSibling;
  details.classList.toggle('visible');
}
function toggleCategory(el) {
  el.classList.toggle('collapsed');
  var body = el.nextElementSibling;
  body.classList.toggle('collapsed');
}
function expandAll() {
  document.querySelectorAll('.check-row-summary').forEach(function(el){ el.classList.add('expanded'); el.nextElementSibling.classList.add('visible'); });
  document.querySelectorAll('.category-header').forEach(function(el){ el.classList.remove('collapsed'); el.nextElementSibling.classList.remove('collapsed'); });
}
function collapseAll() {
  document.querySelectorAll('.check-row-summary').forEach(function(el){ el.classList.remove('expanded'); el.nextElementSibling.classList.remove('visible'); });
}
</script>

</body>
</html>
"@

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    return $OutputPath
}

function New-DiagnosticJsonReport {
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][object]$Analysis,
        [Parameter(Mandatory)][hashtable]$Environment,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $exportObject = [PSCustomObject]@{
        GeneratedAt = Get-Date -Format 'o'
        Environment = $Environment
        HealthScore = $Analysis.HealthScore
        Counts      = $Analysis.Counts
        RootCause   = $Analysis.RootCause.Hypotheses
        EngineerSummary = $Analysis.EngineerSummary
        Checks      = $Results
    }

    $exportObject | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
    return $OutputPath
}
