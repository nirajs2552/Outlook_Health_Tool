#requires -Version 5.1
<#
.SYNOPSIS
    SystemChecks.ps1 - Windows/system-level diagnostics (Section 8 of spec).
    Read-only. Every function returns one or more New-CheckResult objects.
#>

function Get-SystemInfoSummary {
    # Used by the GUI header (User / Computer / OS) - not a scored check.
    Invoke-SafeCheck -Category 'System' -Check 'System Summary' -ScriptBlock {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            UserName     = "$env:USERDOMAIN\$env:USERNAME"
            OSCaption    = $os.Caption
            OSBuild      = $os.BuildNumber
            Domain       = $cs.Domain
        }
    }
}

function Test-SystemDiskSpace {
    Invoke-SafeCheck -Category 'System' -Check 'Disk Free Space' -ScriptBlock {
        $cmd = "Get-CimInstance Win32_LogicalDisk -Filter `"DeviceID='$env:SystemDrive'`""
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'" -ErrorAction Stop
        $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
        $totalGB = [math]::Round($disk.Size / 1GB, 1)
        $pctFree = if ($disk.Size -gt 0) { [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1) } else { 0 }

        $status = 'PASS'; $rec = ''
        if ($freeGB -lt 2) {
            $status = 'CRITICAL'
            $rec = 'System drive is nearly full. Free up space immediately - Outlook, Windows Update, and Office repair can all fail or corrupt data files below ~2 GB free.'
        } elseif ($freeGB -lt 10) {
            $status = 'WARNING'
            $rec = 'Free disk space is low. Low disk space is a common cause of OST corruption, failed updates, and Outlook hangs.'
        }

        New-CheckResult -Category 'System' -Check 'Disk Free Space' -Status $status `
            -Value "$freeGB GB free of $totalGB GB ($pctFree% free) on $env:SystemDrive" `
            -Expected 'At least 10 GB free on the system drive' `
            -Finding "$(if($status -eq 'PASS'){'Sufficient free space.'}else{'Free space is below the healthy threshold.'})" `
            -Recommendation $rec -Command $cmd -RawResult ($disk | Out-String)
    }
}

function Test-SystemUptime {
    Invoke-SafeCheck -Category 'System' -Check 'System Uptime' -ScriptBlock {
        $cmd = 'Get-CimInstance Win32_OperatingSystem | Select LastBootUpTime'
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $lastBoot = $os.LastBootUpTime
        $uptime = (Get-Date) - $lastBoot
        $days = [math]::Round($uptime.TotalDays, 1)

        $status = 'PASS'; $rec = ''
        if ($days -ge 14) {
            $status = 'WARNING'
            $rec = 'The system has not been restarted in over two weeks. Long uptime can contribute to memory fragmentation, stuck update installs, and general instability. A restart is a low-risk troubleshooting step.'
        }

        New-CheckResult -Category 'System' -Check 'System Uptime' -Status $status `
            -Value "$days days (last boot: $lastBoot)" -Expected 'Restarted within the last 14 days' `
            -Finding "$(if($status -eq 'PASS'){'Uptime is within a healthy range.'}else{'Uptime is longer than typically recommended.'})" `
            -Recommendation $rec -Command $cmd -RawResult "LastBootUpTime: $lastBoot"
    }
}

function Test-SystemTimeSync {
    Invoke-SafeCheck -Category 'System' -Check 'Time Synchronization' -ScriptBlock {
        $cmd = 'w32tm /query /status'
        $status = 'INFO'; $value = 'Unable to determine'; $finding = 'Time service status could not be fully verified.'
        $rec = ''
        try {
            $raw = & w32tm /query /status 2>&1 | Out-String
            if ($raw -match 'Leap Indicator') {
                $offsetMatch = [regex]::Match($raw, 'Last Successful Sync Time:\s*(.+)')
                $value = if ($offsetMatch.Success) { "Last sync: $($offsetMatch.Groups[1].Value.Trim())" } else { 'Time service responding' }
                $status = 'PASS'
                $finding = 'Windows Time service is running and reporting status.'
            } else {
                $status = 'WARNING'
                $finding = 'Windows Time service did not return a normal status.'
                $rec = 'Incorrect system time can break TLS/Autodiscover/Exchange Online authentication (Kerberos and OAuth are time-sensitive). Verify the clock and time zone.'
            }
        } catch {
            $status = 'WARNING'
            $finding = 'Could not query the Windows Time service.'
            $rec = 'Verify system time manually; large clock skew breaks Microsoft 365 authentication.'
        }
        New-CheckResult -Category 'System' -Check 'Time Synchronization' -Status $status `
            -Value $value -Expected 'System clock in sync with a reliable time source' `
            -Finding $finding -Recommendation $rec -Command $cmd -RawResult $raw
    }
}

function Get-SystemNetworkAdapterInfo {
    Invoke-SafeCheck -Category 'System' -Check 'Network Adapters' -ScriptBlock {
        $cmd = 'Get-NetIPConfiguration'
        $status = 'PASS'; $rec = ''
        try {
            $ipconfig = Get-NetIPConfiguration -ErrorAction Stop | Where-Object { $_.IPv4Address }
            if (-not $ipconfig) {
                $status = 'CRITICAL'
                $rec = 'No active network adapter with an IPv4 address was found. Outlook cannot connect to Exchange Online without network connectivity.'
                $value = 'No active adapters with an IPv4 address'
            } else {
                $adapterSummary = ($ipconfig | ForEach-Object {
                    "$($_.InterfaceAlias): $($_.IPv4Address.IPAddress) (GW: $($_.IPv4DefaultGateway.NextHop))"
                }) -join '; '
                $value = $adapterSummary
            }
        } catch {
            $status = 'ERROR'
            $value = 'Unable to enumerate adapters'
        }
        New-CheckResult -Category 'System' -Check 'Network Adapters' -Status $status `
            -Value $value -Expected 'At least one active adapter with a valid IPv4 address and gateway' `
            -Finding "$(if($status -eq 'PASS'){'Active network adapter(s) detected.'}else{'No usable network adapter detected.'})" `
            -Recommendation $rec -Command $cmd -RawResult ($ipconfig | Out-String)
    }
}

function Get-SystemHardwareInfo {
    Invoke-SafeCheck -Category 'System' -Check 'Hardware / OS Summary' -ScriptBlock {
        $cmd = 'Get-CimInstance Win32_OperatingSystem, Win32_ComputerSystem, Win32_Processor'
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        $ramGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)

        $value = "$($os.Caption) (Build $($os.BuildNumber)), $($cs.SystemType), $ramGB GB RAM, PowerShell $($PSVersionTable.PSVersion)"

        New-CheckResult -Category 'System' -Check 'Hardware / OS Summary' -Status 'INFO' `
            -Value $value -Expected 'Informational' `
            -Finding 'Baseline system information collected for correlation with other findings.' `
            -Command $cmd -RawResult ($os | Out-String)
    }
}

function Test-SystemUserPermissions {
    Invoke-SafeCheck -Category 'System' -Check 'User Permissions' -ScriptBlock {
        $cmd = '[Security.Principal.WindowsPrincipal]::new(...).IsInRole(Administrator)'
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        New-CheckResult -Category 'System' -Check 'User Permissions' -Status 'INFO' `
            -Value "Running as administrator: $isAdmin" -Expected 'Informational (standard user context is normal and expected)' `
            -Finding "This tool is running with $(if($isAdmin){'administrator'}else{'standard user'}) privileges. Some diagnostics (deep event log access, certain services) may be limited without admin rights." `
            -Command $cmd
    }
}

function Test-RelevantWindowsServices {
    Invoke-SafeCheck -Category 'System' -Check 'Relevant Windows Services' -ScriptBlock {
        $cmd = 'Get-Service ClickToRunSvc, VaultSvc, TokenBroker, CryptSvc, BITS'
        $serviceMap = @{
            'ClickToRunSvc' = 'Microsoft Office Click-to-Run'
            'VaultSvc'      = 'Credential Manager'
            'TokenBroker'   = 'Web Account Manager'
            'CryptSvc'      = 'Cryptographic Services'
            'BITS'          = 'Background Intelligent Transfer Service'
        }

        $rows = foreach ($key in $serviceMap.Keys) {
            $svc = Get-Service -Name $key -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                Service     = $serviceMap[$key]
                Status      = if ($svc) { $svc.Status } else { 'Not found' }
                StartType   = if ($svc) { $svc.StartType } else { 'N/A' }
            }
        }

        # Only flag services that are genuinely stopped when they matter (Automatic and stopped).
        $problems = $rows | Where-Object { $_.Status -eq 'Stopped' -and $_.StartType -eq 'Automatic' }
        $status = if ($problems.Count -gt 0) { 'WARNING' } else { 'PASS' }
        $rec = if ($problems.Count -gt 0) {
            "The following service(s) are set to start automatically but are currently stopped: $(($problems.Service) -join ', '). This can affect Office activation, authentication, or updates."
        } else { '' }

        New-CheckResult -Category 'System' -Check 'Relevant Windows Services' -Status $status `
            -Value (($rows | ForEach-Object { "$($_.Service): $($_.Status)" }) -join ' | ') `
            -Expected 'Relevant services running or available on demand' `
            -Finding "$(if($status -eq 'PASS'){'Relevant Windows services are in an expected state.'}else{'One or more relevant services are stopped when they should be running.'})" `
            -Recommendation $rec -Command $cmd -RawResult ($rows | Format-Table -AutoSize | Out-String)
    }
}

function Invoke-AllSystemChecks {
    @(
        Test-SystemDiskSpace
        Test-SystemUptime
        Test-SystemTimeSync
        Get-SystemNetworkAdapterInfo
        Get-SystemHardwareInfo
        Test-SystemUserPermissions
        Test-RelevantWindowsServices
    )
}
