#requires -Version 5.1
<#
.SYNOPSIS
    OutlookChecks.ps1 - Outlook installation, process, profile, data-file,
    cached-mode, add-in, and logging diagnostics. Read-only.
#>

# ----------------------------------------------------------------------------
# 9. Outlook installation / version detection (Classic vs New Outlook)
# ----------------------------------------------------------------------------
function Get-OutlookInstallationInfo {
    Invoke-SafeCheck -Category 'Outlook' -Check 'Outlook Installation' -ScriptBlock {
        $cmd = 'HKLM/HKCU:\...\Office\ClickToRun\Configuration ; Get-Item OUTLOOK.EXE VersionInfo'
        $exePath = $null
        foreach ($p in @(
            "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
            "$env:ProgramFiles\Microsoft Office\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE"
        )) {
            if (Test-Path -LiteralPath $p) { $exePath = $p; break }
        }

        $ctrConfig = Get-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        $isC2R = $null -ne $ctrConfig
        $productReleaseIds = if ($ctrConfig) { $ctrConfig.ProductReleaseIds } else { $null }

        if (-not $exePath) {
            return New-CheckResult -Category 'Outlook' -Check 'Outlook Installation' -Status 'CRITICAL' `
                -Value 'OUTLOOK.EXE not found in standard install paths' `
                -Expected 'Outlook executable present under Program Files' `
                -Finding 'No Outlook installation could be located on this machine.' `
                -Recommendation 'Confirm Microsoft 365 Apps / Office is installed. Reinstall Office if it should be present.' `
                -Command $cmd
        }

        $verInfo = (Get-Item -LiteralPath $exePath).VersionInfo
        $fileVersion = $verInfo.ProductVersion

        $installType = if ($isC2R) { 'Click-to-Run (Microsoft 365 Apps / modern MSI-free install)' } else { 'MSI-based install' }

        New-CheckResult -Category 'Outlook' -Check 'Outlook Installation' -Status 'PASS' `
            -Value "$exePath | Version $fileVersion | $installType$(if($productReleaseIds){" | $productReleaseIds"})" `
            -Expected 'A supported Outlook version (2016/2019/2021/Microsoft 365 Apps) installed' `
            -Finding 'Outlook installation detected.' `
            -Command $cmd -RawResult ($verInfo | Out-String)
    }
}

function Get-NewOutlookStatus {
    Invoke-SafeCheck -Category 'Outlook' -Check 'Classic vs New Outlook' -ScriptBlock {
        $cmd = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General (NewOutlookInstalled) + package check'
        $newOutlookPkg = Get-AppxPackage -Name 'Microsoft.OutlookForWindows' -ErrorAction SilentlyContinue
        $newOutlookInstalled = $null -ne $newOutlookPkg

        $regFlag = Get-RegistryValueSafe -Path 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General' -Name 'NewOutlookInstalled'

        $value = if ($newOutlookInstalled) {
            "New Outlook (UWP) package installed: $($newOutlookPkg.Version)"
        } else {
            'New Outlook (UWP) package not detected - Classic Outlook only'
        }

        New-CheckResult -Category 'Outlook' -Check 'Classic vs New Outlook' -Status 'INFO' `
            -Value $value -Expected 'Informational - identifies which Outlook client is in use' `
            -Finding 'Client type identified so remaining checks can be interpreted correctly. Note: many registry-based diagnostics below (profiles, OST, add-ins, cached mode) apply to Classic Outlook; New Outlook stores configuration differently and has limited local diagnosability.' `
            -Command $cmd -RawResult "AppxPackage: $($newOutlookPkg -ne $null) | RegFlag: $regFlag"
    }
}

# ----------------------------------------------------------------------------
# 10. Outlook process checks
# ----------------------------------------------------------------------------
function Test-OutlookProcess {
    Invoke-SafeCheck -Category 'Outlook' -Check 'Outlook Process' -ScriptBlock {
        $cmd = 'Get-Process OUTLOOK -ErrorAction SilentlyContinue'
        $procs = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue

        if (-not $procs) {
            return New-CheckResult -Category 'Outlook' -Check 'Outlook Process' -Status 'INFO' `
                -Value 'OUTLOOK.EXE is not currently running' -Expected 'Informational' `
                -Finding 'Outlook is not running. This is not inherently a problem - only relevant if the user reports Outlook is currently unresponsive or won''t stay open.' `
                -Command $cmd
        }

        $proc = $procs | Select-Object -First 1
        $runningDays = [math]::Round(((Get-Date) - $proc.StartTime).TotalDays, 1)
        $memMB = [math]::Round($proc.WorkingSet64 / 1MB, 0)
        $responding = $proc.Responding

        $status = 'PASS'; $rec = ''
        if (-not $responding) {
            $status = 'CRITICAL'
            $rec = 'Outlook is not responding (hung). Consider restarting Outlook; if it recurs, check add-ins and OST health.'
        } elseif ($runningDays -ge 14) {
            $status = 'WARNING'
            $rec = 'Outlook has been running continuously for an extended period without restart. Long-running sessions are more prone to memory bloat and instability - a restart is a safe first step.'
        } elseif ($procs.Count -gt 1) {
            $status = 'WARNING'
            $rec = 'Multiple OUTLOOK.EXE processes detected, which can indicate a hung instance or an add-in spawning extra processes.'
        }

        New-CheckResult -Category 'Outlook' -Check 'Outlook Process' -Status $status `
            -Value "Running $($procs.Count) process(es) | Started: $($proc.StartTime) ($runningDays days ago) | Memory: $memMB MB | Responding: $responding" `
            -Expected 'Single responsive Outlook process' `
            -Finding "$(if($status -eq 'PASS'){'Outlook process is healthy and responding.'}else{'Outlook process shows a potential issue.'})" `
            -Recommendation $rec -Command $cmd -RawResult ($procs | Format-Table Id, StartTime, Responding, @{n='MemMB';e={[math]::Round($_.WorkingSet64/1MB,0)}} | Out-String)
    }
}

# ----------------------------------------------------------------------------
# 11. Outlook profile checks
# ----------------------------------------------------------------------------
function Test-OutlookProfiles {
    Invoke-SafeCheck -Category 'Outlook Profile' -Check 'Mail Profiles' -ScriptBlock {
        $cmd = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
        $profilesPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
        $profiles = Get-RegistryChildrenSafe -Path $profilesPath

        if (-not $profiles -or $profiles.Count -eq 0) {
            return New-CheckResult -Category 'Outlook Profile' -Check 'Mail Profiles' -Status 'CRITICAL' `
                -Value 'No mail profiles found' -Expected 'At least one valid Outlook mail profile' `
                -Finding 'No Outlook mail profile exists under the current Windows Messaging Subsystem key. Note: New Outlook does not use this legacy profile store, so this finding does not apply if New Outlook is the active client.' `
                -Recommendation 'Create a new Outlook profile (Control Panel > Mail > Show Profiles > Add), or confirm whether New Outlook is in use.' `
                -Command $cmd
        }

        $profileNames = $profiles | ForEach-Object { $_.PSChildName }
        $status = 'PASS'
        $finding = "Found $($profiles.Count) profile(s): $($profileNames -join ', ')"

        New-CheckResult -Category 'Outlook Profile' -Check 'Mail Profiles' -Status $status `
            -Value "$($profiles.Count) profile(s) - $($profileNames -join ', ')" `
            -Expected 'At least one valid mail profile' -Finding $finding `
            -Command $cmd -RawResult ($profileNames -join "`n")
    }
}

function Test-OutlookDefaultProfile {
    Invoke-SafeCheck -Category 'Outlook Profile' -Check 'Default Mail Profile' -ScriptBlock {
        $cmd = 'HKCU:\Software\Microsoft\Office\16.0\Outlook (DefaultProfile)'
        $ver = Get-OutlookOfficeVersionKey
        $defaultProfile = Get-RegistryValueSafe -Path "HKCU:\Software\Microsoft\Office\$ver\Outlook" -Name 'DefaultProfile'

        $profilesPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
        $existingProfiles = (Get-RegistryChildrenSafe -Path $profilesPath) | ForEach-Object { $_.PSChildName }

        if (-not $defaultProfile) {
            return New-CheckResult -Category 'Outlook Profile' -Check 'Default Mail Profile' -Status 'WARNING' `
                -Value 'No default profile configured' -Expected 'A default profile name set' `
                -Finding 'Outlook has no explicit default profile set. Outlook may prompt to choose a profile at startup, or this may be normal if "always use this profile" is not configured.' `
                -Recommendation 'If the user reports being prompted to pick a profile, set a default profile in Control Panel > Mail.' `
                -Command $cmd
        }

        $status = 'PASS'; $rec = ''
        if ($existingProfiles -and $defaultProfile -notin $existingProfiles) {
            $status = 'CRITICAL'
            $rec = 'The configured default profile does not match any existing profile. This will cause Outlook to fail to start correctly or prompt unexpectedly. Recreate or reselect the default profile.'
        }

        New-CheckResult -Category 'Outlook Profile' -Check 'Default Mail Profile' -Status $status `
            -Value "Default profile: $defaultProfile" -Expected 'Default profile name matches an existing profile' `
            -Finding "$(if($status -eq 'PASS'){'Default profile is configured and matches an existing profile.'}else{'Default profile is misconfigured.'})" `
            -Recommendation $rec -Command $cmd
    }
}

# ----------------------------------------------------------------------------
# 12/13. OST/PST + Cached Exchange Mode
# ----------------------------------------------------------------------------
function Test-OutlookDataFiles {
    Invoke-SafeCheck -Category 'OST/PST' -Check 'Outlook Data Files (OST/PST)' -ScriptBlock {
        $cmd = 'Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Outlook" -Include *.ost,*.pst -Recurse'
        $outlookDataPath = "$env:LOCALAPPDATA\Microsoft\Outlook"
        $results = @()

        if (-not (Test-Path -LiteralPath $outlookDataPath)) {
            return New-CheckResult -Category 'OST/PST' -Check 'Outlook Data Files (OST/PST)' -Status 'CRITICAL' `
                -Value "Default Outlook data folder not found: $outlookDataPath" `
                -Expected 'Outlook data folder exists with at least one OST/PST file' `
                -Finding 'The default Outlook data file folder does not exist on this profile.' `
                -Recommendation 'This suggests Outlook has never been fully configured for this Windows user, or the profile points elsewhere. Verify profile configuration.' `
                -Command $cmd
        }

        $files = Get-ChildItem -LiteralPath $outlookDataPath -Include *.ost, *.pst -Recurse -ErrorAction SilentlyContinue

        if (-not $files -or $files.Count -eq 0) {
            return New-CheckResult -Category 'OST/PST' -Check 'Outlook Data Files (OST/PST)' -Status 'CRITICAL' `
                -Value "No .ost or .pst files found under $outlookDataPath" `
                -Expected 'At least one OST file matching the configured profile' `
                -Finding 'Outlook data folder exists, but no data files were found inside it.' `
                -Recommendation 'The OST may need to be recreated, or the mailbox has never fully synced. Do not delete/recreate automatically - confirm with the user first.' `
                -Command $cmd
        }

        $summaries = foreach ($f in $files) {
            $sizeGB = [math]::Round($f.Length / 1GB, 2)
            $ageHours = [math]::Round(((Get-Date) - $f.LastWriteTime).TotalHours, 1)
            "$($f.Name): $sizeGB GB, last written $ageHours hrs ago"
        }

        $status = 'PASS'; $rec = ''
        $largeFiles = $files | Where-Object { $_.Length -gt 40GB }
        $staleFiles = $files | Where-Object { ((Get-Date) - $_.LastWriteTime).TotalDays -gt 3 }

        if ($largeFiles) {
            $status = 'WARNING'
            $rec = 'One or more data files exceed 40 GB. Very large OST files are more prone to performance issues and (rarely) corruption; consider reducing the cached mail range.'
        } elseif ($staleFiles) {
            $status = 'WARNING'
            $rec = 'A data file has not been written to in over 3 days, which may indicate sync has stalled. Confirm Outlook is actively syncing.'
        }

        New-CheckResult -Category 'OST/PST' -Check 'Outlook Data Files (OST/PST)' -Status $status `
            -Value ($summaries -join ' | ') -Expected 'OST file present, reasonably sized, and recently active' `
            -Finding "$(if($status -eq 'PASS'){'Data file(s) present with recent activity.'}else{'Data file(s) show a potential concern.'})" `
            -Recommendation $rec -Command $cmd -RawResult ($files | Format-Table Name, @{n='SizeGB';e={[math]::Round($_.Length/1GB,2)}}, LastWriteTime | Out-String)
    }
}

function Test-OutlookCachedMode {
    Invoke-SafeCheck -Category 'Outlook' -Check 'Cached Exchange Mode' -ScriptBlock {
        $ver = Get-OutlookOfficeVersionKey
        $cmd = "HKCU:\Software\Microsoft\Office\$ver\Outlook\Cached Mode"
        $base = "HKCU:\Software\Microsoft\Office\$ver\Outlook\Cached Mode"
        $enabled = Get-RegistryValueSafe -Path $base -Name 'Enable'
        $syncWindowSetting = Get-RegistryValueSafe -Path $base -Name 'SyncWindowSetting'

        $monthsMap = @{ 0 = 'All'; 1 = '3 months'; 3 = '6 months'; 6 = '12 months'; 12 = '24 months'; 24 = '36 months' }
        $monthsText = if ($null -ne $syncWindowSetting -and $monthsMap.ContainsKey([int]$syncWindowSetting)) {
            $monthsMap[[int]$syncWindowSetting]
        } else { 'Default/Unspecified' }

        if ($null -eq $enabled) {
            return New-CheckResult -Category 'Outlook' -Check 'Cached Exchange Mode' -Status 'INFO' `
                -Value 'No explicit Cached Mode registry value found (likely using Outlook default: enabled)' `
                -Expected 'Cached Exchange Mode enabled for Exchange Online mailboxes' `
                -Finding 'Cached mode setting not explicitly configured; Outlook''s default is enabled.' `
                -Command $cmd
        }

        $status = if ($enabled -eq 1) { 'PASS' } else { 'WARNING' }
        $rec = if ($enabled -ne 1) {
            'Cached Exchange Mode is disabled. In an Exchange Online environment, running in online-only mode is more sensitive to network latency and can contribute to slowness or timeout-related symptoms.'
        } else { '' }

        New-CheckResult -Category 'Outlook' -Check 'Cached Exchange Mode' -Status $status `
            -Value "Enabled: $($enabled -eq 1) | Sync window: $monthsText" `
            -Expected 'Cached Exchange Mode enabled' `
            -Finding "$(if($status -eq 'PASS'){'Cached Exchange Mode is enabled.'}else{'Cached Exchange Mode is disabled.'})" `
            -Recommendation $rec -Command $cmd
    }
}

# ----------------------------------------------------------------------------
# 14. Add-in checks (COM add-ins, load behavior)
# ----------------------------------------------------------------------------
function Test-OutlookAddins {
    Invoke-SafeCheck -Category 'Add-ins' -Check 'COM / Outlook Add-ins' -ScriptBlock {
        $cmd = 'HKCU:\Software\Microsoft\Office\Outlook\Addins ; HKLM:\Software\Microsoft\Office\Outlook\Addins'
        $paths = @(
            'HKCU:\Software\Microsoft\Office\Outlook\Addins',
            'HKLM:\Software\Microsoft\Office\Outlook\Addins',
            'HKLM:\Software\WOW6432Node\Microsoft\Office\Outlook\Addins'
        )

        # LoadBehavior: 3 = load at startup (enabled). 0/2/8/16 generally mean not loading.
        $addins = @()
        foreach ($p in $paths) {
            $children = Get-RegistryChildrenSafe -Path $p
            foreach ($c in $children) {
                $props = Get-RegistryValueSafe -Path $c.PSPath
                $loadBehavior = $props.LoadBehavior
                $friendlyName = $props.FriendlyName
                $addins += [PSCustomObject]@{
                    Name         = if ($friendlyName) { $friendlyName } else { $c.PSChildName }
                    ProgId       = $c.PSChildName
                    LoadBehavior = $loadBehavior
                    Hive         = $p
                }
            }
        }

        if ($addins.Count -eq 0) {
            return New-CheckResult -Category 'Add-ins' -Check 'COM / Outlook Add-ins' -Status 'INFO' `
                -Value 'No COM add-ins registered' -Expected 'Informational' `
                -Finding 'No Outlook COM add-ins found in the standard registry locations.' -Command $cmd
        }

        $disabled = $addins | Where-Object { $_.LoadBehavior -ne 3 }
        $enabled = $addins | Where-Object { $_.LoadBehavior -eq 3 }

        $status = if ($disabled.Count -gt 0) { 'WARNING' } else { 'PASS' }
        $rec = if ($disabled.Count -gt 0) {
            'One or more add-ins are not set to load at startup (LoadBehavior != 3). This can be intentional (user-disabled) or can indicate Outlook auto-disabled a crashing add-in - check Outlook''s COM Add-ins dialog for "was disabled due to a problem".'
        } else { '' }

        $value = "Enabled: $($enabled.Count) | Not loading: $($disabled.Count)"
        if ($disabled.Count -gt 0) {
            $value += " -- Not loading: " + (($disabled | ForEach-Object { "$($_.Name) (LoadBehavior=$($_.LoadBehavior))" }) -join '; ')
        }

        New-CheckResult -Category 'Add-ins' -Check 'COM / Outlook Add-ins' -Status $status `
            -Value $value -Expected 'Add-ins load as expected (LoadBehavior=3) unless intentionally disabled' `
            -Finding "$(if($status -eq 'PASS'){'All registered add-ins are set to load normally.'}else{'One or more add-ins are not loading at startup.'})" `
            -Recommendation $rec -Command $cmd -RawResult ($addins | Format-Table Name, ProgId, LoadBehavior | Out-String)
    }
}

# ----------------------------------------------------------------------------
# 29. Safe Mode detection
# ----------------------------------------------------------------------------
function Test-OutlookSafeMode {
    Invoke-SafeCheck -Category 'Outlook' -Check 'Safe Mode Indicator' -ScriptBlock {
        $cmd = 'Get-CimInstance Win32_Process -Filter "Name=''OUTLOOK.EXE''" | Select CommandLine'
        $proc = Get-CimInstance Win32_Process -Filter "Name='OUTLOOK.EXE'" -ErrorAction SilentlyContinue | Select-Object -First 1

        if (-not $proc) {
            return New-CheckResult -Category 'Outlook' -Check 'Safe Mode Indicator' -Status 'INFO' `
                -Value 'Outlook is not running - cannot inspect command line' -Expected 'Informational' `
                -Finding 'No running Outlook process to inspect for safe-mode flags.' -Command $cmd
        }

        $cmdLine = $proc.CommandLine
        $isSafeMode = $cmdLine -match '/safe' -or $cmdLine -match 'safe:1'

        $status = if ($isSafeMode) { 'WARNING' } else { 'PASS' }
        $rec = if ($isSafeMode) { 'Outlook is running in Safe Mode, which disables add-ins and some customizations. This is often a troubleshooting step already taken, or can indicate Outlook detected repeated startup failures.' } else { '' }

        New-CheckResult -Category 'Outlook' -Check 'Safe Mode Indicator' -Status $status `
            -Value "Safe mode: $isSafeMode" -Expected 'Outlook running normally (not in Safe Mode)' `
            -Finding "$(if($status -eq 'PASS'){'Outlook is not running in Safe Mode.'}else{'Outlook is currently running in Safe Mode.'})" `
            -Recommendation $rec -Command $cmd -RawResult (Protect-SensitiveText -Text $cmdLine)
    }
}

# ----------------------------------------------------------------------------
# 28. Outlook diagnostic logging
# ----------------------------------------------------------------------------
function Test-OutlookLogging {
    Invoke-SafeCheck -Category 'Outlook' -Check 'Diagnostic Logging' -ScriptBlock {
        $ver = Get-OutlookOfficeVersionKey
        $cmd = "HKCU:\Software\Microsoft\Office\$ver\Outlook (EnableLogging)"
        $enableLogging = Get-RegistryValueSafe -Path "HKCU:\Software\Microsoft\Office\$ver\Outlook" -Name 'EnableLogging'

        $isEnabled = $enableLogging -eq 1
        New-CheckResult -Category 'Outlook' -Check 'Diagnostic Logging' -Status 'INFO' `
            -Value "Diagnostic logging enabled: $isEnabled" -Expected 'Informational' `
            -Finding "$(if($isEnabled){'Outlook diagnostic logging is currently enabled - useful for deep troubleshooting but adds overhead and disk I/O.'}else{'Outlook diagnostic logging is not enabled.'})" `
            -Command $cmd
    }
}

# ----------------------------------------------------------------------------
# 36. Outlook startup configuration (e.g. /resetnavpane, addin startup switches)
# ----------------------------------------------------------------------------
function Test-OutlookStartupConfig {
    Invoke-SafeCheck -Category 'Outlook' -Check 'Startup Configuration' -ScriptBlock {
        $ver = Get-OutlookOfficeVersionKey
        $cmd = "HKCU:\Software\Microsoft\Office\$ver\Outlook\Options\General ; Run key inspection"
        $runKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
        $runEntries = Get-RegistryValueSafe -Path $runKeyPath
        $outlookAutoStart = $false
        if ($runEntries) {
            $props = $runEntries.PSObject.Properties | Where-Object { $_.Value -match 'OUTLOOK\.EXE' -and $_.Name -notmatch '^PS' }
            $outlookAutoStart = [bool]$props
        }

        New-CheckResult -Category 'Outlook' -Check 'Startup Configuration' -Status 'INFO' `
            -Value "Outlook configured to auto-start with Windows: $outlookAutoStart" -Expected 'Informational' `
            -Finding 'Startup configuration collected for context; not inherently good or bad.' -Command $cmd
    }
}

function Invoke-AllOutlookChecks {
    @(
        Get-OutlookInstallationInfo
        Get-NewOutlookStatus
        Test-OutlookProcess
        Test-OutlookProfiles
        Test-OutlookDefaultProfile
        Test-OutlookDataFiles
        Test-OutlookCachedMode
        Test-OutlookAddins
        Test-OutlookSafeMode
        Test-OutlookLogging
        Test-OutlookStartupConfig
    )
}
