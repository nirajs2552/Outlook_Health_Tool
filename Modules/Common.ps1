#requires -Version 5.1
<##
.SYNOPSIS
    Common.ps1 - Shared helper functions used by every diagnostic module.
#>

$Script:LogFilePath = $null

function Initialize-DiagLog {
    param([Parameter(Mandatory)][string]$LogFolder)
    if (-not (Test-Path -LiteralPath $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd'
    $Script:LogFilePath = Join-Path $LogFolder "OutlookHealthCheck-$stamp.log"
    Write-DiagLog -Message "===== Outlook Health Check session started =====" -Level 'INFO'
}

function Write-DiagLog {
    param([Parameter(Mandatory)][string]$Message,[ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level='INFO')
    if (-not $Script:LogFilePath) { return }
    $line = "{0:yyyy-MM-dd HH:mm:ss}  [{1}]  {2}" -f (Get-Date),$Level,(Protect-SensitiveText -Text $Message)
    try { Add-Content -LiteralPath $Script:LogFilePath -Value $line -Encoding UTF8 -ErrorAction Stop } catch { }
}

function Protect-SensitiveText {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $patterns = @(
        '(?i)(password|pwd|secret|token|refresh_token|access_token|apikey|api_key)\s*[:=]\s*\S+',
        '\b[A-Za-z0-9\-_\.]{40,}\b',
        '\b([A-Z0-9]{5}-){4}[A-Z0-9]{5}\b'
    )
    $result=$Text
    foreach($p in $patterns){ $result=[regex]::Replace($result,$p,'[REDACTED]') }
    return $result
}

function New-CheckResult {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Check,
        [Parameter(Mandatory)][ValidateSet('PASS','INFO','WARNING','CRITICAL','ERROR')][string]$Status,
        [ValidateSet('PASS','INFO','WARNING','CRITICAL','ERROR')][string]$Severity=$Status,
        [string]$Value='', [string]$Expected='', [string]$Finding='', [string]$Recommendation='', [string]$Command='', [string]$RawResult=''
    )
    [PSCustomObject]@{
        Category=$Category; Check=$Check; Status=$Status; Severity=$Severity
        Value=(Protect-SensitiveText -Text $Value); Expected=$Expected
        Finding=(Protect-SensitiveText -Text $Finding); Recommendation=$Recommendation
        Command=$Command; RawResult=(Protect-SensitiveText -Text $RawResult); Timestamp=Get-Date
    }
}

function Invoke-SafeCheck {
    param([Parameter(Mandatory)][string]$Category,[Parameter(Mandatory)][string]$Check,[Parameter(Mandatory)][scriptblock]$ScriptBlock)
    try {
        Write-DiagLog -Message "Running check: [$Category] $Check" -Level 'DEBUG'
        $result=& $ScriptBlock
        if(-not $result){ return New-CheckResult -Category $Category -Check $Check -Status 'ERROR' -Finding 'Diagnostic returned no result.' -Recommendation 'Re-run the health check.' }
        return $result
    } catch {
        $msg=$_.Exception.Message
        Write-DiagLog -Message "Check FAILED: [$Category] $Check - $msg" -Level 'ERROR'
        return New-CheckResult -Category $Category -Check $Check -Status 'ERROR' -Finding "The diagnostic itself failed to run: $msg" -Recommendation 'This indicates a problem running the diagnostic (e.g. missing permission or module), not necessarily a problem with Outlook.'
    }
}

function Get-RegistryValueSafe {
    param([Parameter(Mandatory)][string]$Path,[string]$Name)
    try {
        if(-not(Test-Path -LiteralPath $Path)){ return $null }
        if($Name){ $item=Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop; return $item.$Name }
        else { return Get-ItemProperty -LiteralPath $Path -ErrorAction Stop }
    } catch { return $null }
}
function Get-RegistryChildrenSafe {
    param([Parameter(Mandatory)][string]$Path)
    try { if(-not(Test-Path -LiteralPath $Path)){ return @() }; return Get-ChildItem -LiteralPath $Path -ErrorAction Stop } catch { return @() }
}
function Test-Command { param([Parameter(Mandatory)][string]$Name); return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue) }
function Get-OutlookOfficeVersionKey {
    foreach($v in @('16.0','15.0')){ $path="HKCU:\Software\Microsoft\Office\$v\Outlook"; if(Test-Path -LiteralPath $path){ return $v } }
    return '16.0'
}
