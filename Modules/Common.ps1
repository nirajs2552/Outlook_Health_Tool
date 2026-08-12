#requires -Version 5.1
<#
.SYNOPSIS
    Common.ps1 - Shared helper functions used by every diagnostic module.

    Provides:
      - New-CheckResult      : factory for the structured diagnostic object
      - Invoke-SafeCheck     : wraps a scriptblock so a single failing check
                                never aborts the overall health check
      - Write-DiagLog        : local log file writer (never logs secrets)
      - Protect-SensitiveText : redaction helper for tokens/passwords/keys
      - Get-RegistryValueSafe: registry read that tolerates missing keys
      - Test-Command         : checks whether a cmdlet/exe is available
#>

# ----------------------------------------------------------------------------
# Global script-level state
# ----------------------------------------------------------------------------
$Script:LogFilePath = $null

function Initialize-DiagLog {
    param([Parameter(Mandatory)][string]$LogFolder)

    if (-not (Test-Path -LiteralPath $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }
    $stamp = Get-Date -Format 'yyyyMMdd'
    $Script:LogFilePath = Join-Path $LogFolder "OutlookHealthCheck-$stamp.log"
    Write-DiagLog -Message "===== Outlook Health Check session started =====" -Level 'INFO'
}

function Write-DiagLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    if (-not $Script:LogFilePath) { return }
    $line = "{0:yyyy-MM-dd HH:mm:ss}  [{1}]  {2}" -f (Get-Date), $Level, (Protect-SensitiveText -Text $Message)
    try {
        Add-Content -LiteralPath $Script:LogFilePath -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch {
        # Logging must never break the diagnostic run.
    }
}

# ----------------------------------------------------------------------------
# Redaction - never let secrets reach the log or the report
# ----------------------------------------------------------------------------
function Protect-SensitiveText {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    $patterns = @(
        # password=..., pwd=..., secret=..., token=...
        '(?i)(password|pwd|secret|token|refresh_token|access_token|apikey|api_key)\s*[:=]\s*\S+',
        # long base64-ish blobs typical of tokens/cookies (40+ chars)
        '\b[A-Za-z0-9\-_\.]{40,}\b',
        # product/license key pattern XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
        '\b([A-Z0-9]{5}-){4}[A-Z0-9]{5}\b'
    )
    $result = $Text
    foreach ($p in $patterns) {
        $result = [regex]::Replace($result, $p, '[REDACTED]')
    }
    return $result
}

# ----------------------------------------------------------------------------
# Structured result object - every single check returns exactly this shape
# ----------------------------------------------------------------------------
function New-CheckResult {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Check,
        [Parameter(Mandatory)][ValidateSet('PASS','INFO','WARNING','CRITICAL','ERROR')][string]$Status,
        [ValidateSet('PASS','INFO','WARNING','CRITICAL','ERROR')][string]$Severity = $Status,
        [string]$Value = '',
        [string]$Expected = '',
        [string]$Finding = '',
        [string]$Recommendation = '',
        [string]$Command = '',
        [string]$RawResult = ''
    )

    [PSCustomObject]@{
        Category       = $Category
        Check          = $Check
        Status         = $Status
        Severity       = $Severity
        Value          = (Protect-SensitiveText -Text $Value)
        Expected       = $Expected
        Finding        = (Protect-SensitiveText -Text $Finding)
        Recommendation = $Recommendation
        Command        = $Command
        RawResult      = (Protect-SensitiveText -Text $RawResult)
        Timestamp      = Get-Date
    }
}

# ----------------------------------------------------------------------------
# Safe execution wrapper - a failed diagnostic becomes an ERROR result,
# never an unhandled exception that stops the whole run.
# ----------------------------------------------------------------------------
function Invoke-SafeCheck {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Check,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )

    try {
        Write-DiagLog -Message "Running check: [$Category] $Check" -Level 'DEBUG'
        $result = & $ScriptBlock
        if (-not $result) {
            return New-CheckResult -Category $Category -Check $Check -Status 'ERROR' `
                -Finding 'Diagnostic returned no result.' -Recommendation 'Re-run the health check.'
        }
        return $result
    }
    catch {
        $msg = $_.Exception.Message
        Write-DiagLog -Message "Check FAILED: [$Category] $Check - $msg" -Level 'ERROR'
        return New-CheckResult -Category $Category -Check $Check -Status 'ERROR' `
            -Finding "The diagnostic itself failed to run: $msg" `
            -Recommendation 'This indicates a problem running the diagnostic (e.g. missing permission or module), not necessarily a problem with Outlook.'
    }
}

# ----------------------------------------------------------------------------
# Registry helper that tolerates missing keys/values across 32/64-bit hives
# ----------------------------------------------------------------------------
function Get-RegistryValueSafe {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Name
    )
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return $null }
        if ($Name) {
            $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
            return $item.$Name
        } else {
            return Get-ItemProperty -LiteralPath $Path -ErrorAction Stop
        }
    } catch {
        return $null
    }
}

function Get-RegistryChildrenSafe {
    param([Parameter(Mandatory)][string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return @() }
        return Get-ChildItem -LiteralPath $Path -ErrorAction Stop
    } catch {
        return @()
    }
}

function Test-Command {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

# ----------------------------------------------------------------------------
# Locate the current user's Outlook "root" registry key (16.0 = Office 2016+
# through Microsoft 365 Apps; also probe 15.0 as a legacy fallback).
# ----------------------------------------------------------------------------
function Get-OutlookOfficeVersionKey {
    $candidates = @('16.0', '15.0')
    foreach ($v in $candidates) {
        $path = "HKCU:\Software\Microsoft\Office\$v\Outlook"
        if (Test-Path -LiteralPath $path) { return $v }
    }
    # Default to 16.0 (current for Microsoft 365 Apps / 2019 / 2021) even if
    # the key doesn't exist yet, so downstream checks can report "not found".
    return '16.0'
}
