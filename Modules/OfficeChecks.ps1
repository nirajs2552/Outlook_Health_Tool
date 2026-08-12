#requires -Version 5.1
<#
.SYNOPSIS
    OfficeChecks.ps1 - Office suite version, activation, and update diagnostics.
    Read-only. Never displays license keys.
#>

function Get-OfficeVersionInfo {
    Invoke-SafeCheck -Category 'Office' -Check 'Office Version' -ScriptBlock {
        $cmd = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        $ctr = Get-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'

        if ($ctr) {
            $value = "Click-to-Run | Version: $($ctr.VersionToReport) | Platform: $($ctr.Platform) | Channel: $($ctr.CDNBaseUrl)"
            return New-CheckResult -Category 'Office' -Check 'Office Version' -Status 'PASS' `
                -Value $value -Expected 'A supported, current Office build' `
                -Finding 'Office Click-to-Run installation detected with version information available.' `
                -Command $cmd -RawResult ($ctr | Out-String)
        }

        # Fall back to MSI-based Office detection
        $msiPath = 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\ProductVersion'
        $msiVer = Get-RegistryValueSafe -Path $msiPath -Name 'LastProduct'
        if ($msiVer) {
            return New-CheckResult -Category 'Office' -Check 'Office Version' -Status 'PASS' `
                -Value "MSI-based Office detected | Version: $msiVer" -Expected 'A supported Office build' `
                -Finding 'MSI-based (non-Click-to-Run) Office installation detected.' -Command $cmd
        }

        New-CheckResult -Category 'Office' -Check 'Office Version' -Status 'WARNING' `
            -Value 'Could not determine Office version via ClickToRun or MSI registry paths' `
            -Expected 'Office version identifiable' `
            -Finding 'Office version could not be determined through standard registry locations. Office may use a non-standard installation, or the detection paths may need updating for a newer Office release.' `
            -Recommendation 'Verify Office version manually via File > Office Account in any Office app.' `
            -Command $cmd
    }
}

function Get-OfficeActivationInfo {
    Invoke-SafeCheck -Category 'Office' -Check 'Office Activation' -ScriptBlock {
        $cmd = 'cscript ospp.vbs /dstatus  (or Get-CimInstance SoftwareLicensingProduct fallback)'
        $osppPaths = @(
            "$env:ProgramFiles\Microsoft Office\Office16",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office16"
        )
        $osppScript = $null
        foreach ($p in $osppPaths) {
            $candidate = Join-Path $p 'ospp.vbs'
            if (Test-Path -LiteralPath $candidate) { $osppScript = $candidate; break }
        }

        $rawOutput = ''
        $licensed = $null
        if ($osppScript) {
            try {
                $rawOutput = & cscript.exe //Nologo $osppScript /dstatus 2>&1 | Out-String
                if ($rawOutput -match 'LICENSE STATUS:\s*---LICENSED---') { $licensed = $true }
                elseif ($rawOutput -match 'LICENSE STATUS:') { $licensed = $false }
            } catch {
                $rawOutput = "ospp.vbs execution failed: $($_.Exception.Message)"
            }
        }

        # Redact any license key fragments defensively even though ospp.vbs
        # partially masks keys itself (shows last 5 chars only).
        $rawOutputSafe = Protect-SensitiveText -Text $rawOutput

        if ($null -eq $licensed) {
            return New-CheckResult -Category 'Office' -Check 'Office Activation' -Status 'INFO' `
                -Value 'Activation status could not be determined via ospp.vbs' `
                -Expected 'Office reports as Licensed/Activated' `
                -Finding 'Could not conclusively determine activation state from local tools. This does not necessarily indicate a problem.' `
                -Recommendation 'Verify activation manually: open any Office app > File > Account > Product Information.' `
                -Command $cmd -RawResult $rawOutputSafe
        }

        $status = if ($licensed) { 'PASS' } else { 'CRITICAL' }
        $rec = if (-not $licensed) {
            'Office does not report as licensed/activated. An unlicensed Office install can cause Outlook to run in reduced-functionality mode or block sending mail. Sign in with the Microsoft 365 account or contact the licensing/IT team.'
        } else { '' }

        New-CheckResult -Category 'Office' -Check 'Office Activation' -Status $status `
            -Value "Licensed: $licensed" -Expected 'Office reports as Licensed' `
            -Finding "$(if($licensed){'Office activation appears healthy.'}else{'Office activation appears invalid or incomplete.'})" `
            -Recommendation $rec -Command $cmd -RawResult $rawOutputSafe
    }
}

function Get-OfficeUpdateInfo {
    Invoke-SafeCheck -Category 'Office' -Check 'Office Update Status' -ScriptBlock {
        $cmd = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration (VersionToReport, UpdateChannel)'
        $ctr = Get-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'

        if (-not $ctr) {
            return New-CheckResult -Category 'Office' -Check 'Office Update Status' -Status 'INFO' `
                -Value 'Not a Click-to-Run install, or update info unavailable' -Expected 'Informational' `
                -Finding 'Office update channel/build information is only available for Click-to-Run installs.' -Command $cmd
        }

        $channel = $ctr.CDNBaseUrl
        $version = $ctr.VersionToReport
        $lastUpdateRaw = Get-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name 'LastScenario'

        New-CheckResult -Category 'Office' -Check 'Office Update Status' -Status 'INFO' `
            -Value "Current build: $version | Update source: $channel" `
            -Expected 'Informational - correlate build/update recency against symptom onset date' `
            -Finding 'Office update/channel information collected for timeline correlation with reported symptoms.' `
            -Command $cmd -RawResult ($ctr | Select-Object VersionToReport, CDNBaseUrl, ClientCulture | Out-String)
    }
}

function Invoke-AllOfficeChecks {
    @(
        Get-OfficeVersionInfo
        Get-OfficeActivationInfo
        Get-OfficeUpdateInfo
    )
}
