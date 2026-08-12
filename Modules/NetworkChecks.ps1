#requires -Version 5.1
<#
.SYNOPSIS
    NetworkChecks.ps1 - DNS, connectivity, proxy, TLS, and Autodiscover
    diagnostics against Microsoft 365 endpoints. Read-only, non-intrusive
    (standard DNS resolution and TCP connect tests only - no scanning).
#>

$Script:M365Endpoints = @(
    'outlook.office365.com',
    'outlook.office.com',
    'login.microsoftonline.com',
    'autodiscover-s.outlook.com',
    'graph.microsoft.com'
)

function Test-DnsResolution {
    Invoke-SafeCheck -Category 'Network' -Check 'DNS Resolution' -ScriptBlock {
        $cmd = 'Resolve-DnsName <endpoint>'
        $results = foreach ($ep in $Script:M365Endpoints) {
            try {
                $r = Resolve-DnsName -Name $ep -Type A -ErrorAction Stop
                [PSCustomObject]@{ Endpoint = $ep; Resolved = $true; Address = ($r | Where-Object { $_.Type -eq 'A' } | Select-Object -First 1 -ExpandProperty IPAddress) }
            } catch {
                [PSCustomObject]@{ Endpoint = $ep; Resolved = $false; Address = $null }
            }
        }

        $failed = $results | Where-Object { -not $_.Resolved }
        $status = if ($failed.Count -eq 0) { 'PASS' } elseif ($failed.Count -eq $results.Count) { 'CRITICAL' } else { 'WARNING' }
        $rec = switch ($status) {
            'CRITICAL' { 'DNS resolution is failing for all tested Microsoft 365 endpoints. Check the DNS server configuration, VPN state, and general internet connectivity.' }
            'WARNING'  { "DNS resolution failed for: $($failed.Endpoint -join ', '). This can indicate a DNS caching issue, split-tunnel VPN misconfiguration, or a filtered/blocked endpoint." }
            default    { '' }
        }

        New-CheckResult -Category 'Network' -Check 'DNS Resolution' -Status $status `
            -Value ("$($results.Count - $failed.Count)/$($results.Count) endpoints resolved") `
            -Expected 'All Microsoft 365 endpoints resolve via DNS' `
            -Finding "$(if($status -eq 'PASS'){'All tested Microsoft 365 endpoints resolved successfully.'}else{'One or more Microsoft 365 endpoints failed to resolve.'})" `
            -Recommendation $rec -Command $cmd -RawResult ($results | Format-Table -AutoSize | Out-String)
    }
}

function Test-M365Connectivity {
    Invoke-SafeCheck -Category 'Network' -Check 'Microsoft 365 HTTPS Connectivity' -ScriptBlock {
        $cmd = 'Test-NetConnection <endpoint> -Port 443'
        $results = foreach ($ep in $Script:M365Endpoints) {
            try {
                $t = Test-NetConnection -ComputerName $ep -Port 443 -WarningAction SilentlyContinue -ErrorAction Stop
                [PSCustomObject]@{ Endpoint = $ep; TcpSucceeded = $t.TcpTestSucceeded }
            } catch {
                [PSCustomObject]@{ Endpoint = $ep; TcpSucceeded = $false }
            }
        }

        $failed = $results | Where-Object { -not $_.TcpSucceeded }
        $status = if ($failed.Count -eq 0) { 'PASS' } elseif ($failed.Count -eq $results.Count) { 'CRITICAL' } else { 'WARNING' }
        $rec = switch ($status) {
            'CRITICAL' { 'TCP 443 connectivity to Microsoft 365 endpoints is failing entirely. Check firewall rules, proxy configuration, and whether a VPN is required/interfering.' }
            'WARNING'  { "TCP 443 connectivity failed for: $($failed.Endpoint -join ', '). This may point to selective firewall/proxy blocking of specific Microsoft 365 services." }
            default    { '' }
        }

        New-CheckResult -Category 'Network' -Check 'Microsoft 365 HTTPS Connectivity' -Status $status `
            -Value ("$($results.Count - $failed.Count)/$($results.Count) endpoints reachable on TCP 443") `
            -Expected 'All Microsoft 365 endpoints reachable on TCP 443' `
            -Finding "$(if($status -eq 'PASS'){'HTTPS (TCP 443) connectivity to Microsoft 365 endpoints is healthy.'}else{'HTTPS connectivity issues detected to one or more Microsoft 365 endpoints.'})" `
            -Recommendation $rec -Command $cmd -RawResult ($results | Format-Table -AutoSize | Out-String)
    }
}

function Test-AutodiscoverConfig {
    Invoke-SafeCheck -Category 'Network' -Check 'Autodiscover' -ScriptBlock {
        $cmd = 'Resolve-DnsName autodiscover-s.outlook.com ; Resolve-DnsName autodiscover.<domain>'
        $coreResolved = $false
        try { $r = Resolve-DnsName -Name 'autodiscover-s.outlook.com' -ErrorAction Stop; $coreResolved = [bool]$r } catch {}

        # Try to derive the user's email domain from the identity registry, without exposing credentials.
        $ver = Get-OutlookOfficeVersionKey
        $identitiesPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities'
        $domain = $null
        $ids = Get-RegistryChildrenSafe -Path $identitiesPath
        foreach ($id in $ids) {
            $props = Get-RegistryValueSafe -Path $id.PSPath
            if ($props.EmailAddress -and $props.EmailAddress -match '@(.+)$') { $domain = $Matches[1]; break }
        }

        $domainResolved = $null
        if ($domain) {
            try { $r2 = Resolve-DnsName -Name "autodiscover.$domain" -ErrorAction Stop; $domainResolved = $true } catch { $domainResolved = $false }
        }

        $status = if ($coreResolved) { 'PASS' } else { 'CRITICAL' }
        $rec = if (-not $coreResolved) { 'The core Microsoft 365 Autodiscover endpoint (autodiscover-s.outlook.com) failed to resolve. Outlook relies on this for mailbox configuration; check DNS and connectivity.' } else { '' }

        $value = "autodiscover-s.outlook.com resolved: $coreResolved"
        if ($domain) { $value += " | autodiscover.$domain resolved: $domainResolved" }

        New-CheckResult -Category 'Network' -Check 'Autodiscover' -Status $status `
            -Value $value -Expected 'Autodiscover endpoints resolve successfully' `
            -Finding "$(if($status -eq 'PASS'){'Autodiscover DNS resolution succeeded.'}else{'Autodiscover DNS resolution failed.'})" `
            -Recommendation $rec -Command $cmd
    }
}

function Get-ProxyConfiguration {
    Invoke-SafeCheck -Category 'Network' -Check 'Proxy Configuration' -ScriptBlock {
        $cmd = 'netsh winhttp show proxy'
        $raw = & netsh.exe winhttp show proxy 2>&1 | Out-String

        $hasProxy = $raw -match 'Proxy Server\(s\)\s*:\s*(\S+)'
        $status = 'PASS'; $rec = ''
        $value = 'No WinHTTP proxy configured (Direct access)'

        if ($hasProxy) {
            $proxyServer = $Matches[1]
            $status = 'INFO'
            $value = "WinHTTP proxy configured: $proxyServer"
            $rec = 'A system-wide proxy is configured. Confirm the proxy allows Microsoft 365 traffic and is not intercepting/breaking TLS for Outlook connections.'
        }

        New-CheckResult -Category 'Network' -Check 'Proxy Configuration' -Status $status `
            -Value $value -Expected 'Direct connection, or a working proxy that permits Microsoft 365 traffic' `
            -Finding "$(if($hasProxy){'A WinHTTP proxy is configured on this machine.'}else{'No system-level WinHTTP proxy is configured.'})" `
            -Recommendation $rec -Command $cmd -RawResult $raw
    }
}

function Test-TlsConfiguration {
    Invoke-SafeCheck -Category 'Network' -Check 'TLS Configuration' -ScriptBlock {
        $cmd = 'HKLM:\...\SCHANNEL\Protocols\TLS 1.2\Client ; [Net.ServicePointManager]::SecurityProtocol'
        $tls12ClientPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client'
        $disabledByDefault = Get-RegistryValueSafe -Path $tls12ClientPath -Name 'DisabledByDefault'
        $enabled = Get-RegistryValueSafe -Path $tls12ClientPath -Name 'Enabled'

        # If no explicit key exists, TLS 1.2 is enabled by default on modern Windows.
        $tls12Explicitlydisabled = ($disabledByDefault -eq 1) -or ($enabled -eq 0)

        $status = if ($tls12Explicitlydisabled) { 'CRITICAL' } else { 'PASS' }
        $rec = if ($tls12Explicitlydisabled) { 'TLS 1.2 appears explicitly disabled for client connections. Microsoft 365 requires TLS 1.2+; this will break Outlook/Exchange Online connectivity. Re-enable TLS 1.2.' } else { '' }

        New-CheckResult -Category 'Network' -Check 'TLS Configuration' -Status $status `
            -Value "TLS 1.2 explicitly disabled: $tls12Explicitlydisabled" `
            -Expected 'TLS 1.2 (or higher) enabled and not disabled by policy' `
            -Finding "$(if($status -eq 'PASS'){'No registry indication that TLS 1.2 is disabled.'}else{'TLS 1.2 appears to be disabled via registry policy.'})" `
            -Recommendation $rec -Command $cmd
    }
}

function Invoke-AllNetworkChecks {
    @(
        Test-DnsResolution
        Test-M365Connectivity
        Test-AutodiscoverConfig
        Get-ProxyConfiguration
        Test-TlsConfiguration
    )
}
