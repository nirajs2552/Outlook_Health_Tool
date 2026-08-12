#requires -Version 5.1
<#
.SYNOPSIS
    IdentityChecks.ps1 - Microsoft 365 identity, Web Account Manager, and
    Credential Manager diagnostics. NEVER collects passwords, tokens, or
    secrets - only entry names/targets and metadata.
#>

function Get-M365IdentityInfo {
    Invoke-SafeCheck -Category 'Authentication' -Check 'Microsoft 365 Identity' -ScriptBlock {
        $cmd = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities'
        $identitiesPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities'
        $identities = Get-RegistryChildrenSafe -Path $identitiesPath

        if (-not $identities -or $identities.Count -eq 0) {
            return New-CheckResult -Category 'Authentication' -Check 'Microsoft 365 Identity' -Status 'WARNING' `
                -Value 'No Office identity entries found' -Expected 'A signed-in work/school (Microsoft 365) identity' `
                -Finding 'No Microsoft 365 / work account identity was found registered with Office on this machine.' `
                -Recommendation 'Sign in to Office with the Microsoft 365 account (any Office app > File > Account > Sign In).' `
                -Command $cmd
        }

        $entries = foreach ($id in $identities) {
            $props = Get-RegistryValueSafe -Path $id.PSPath
            $email = $props.EmailAddress
            $emailMasked = if ($email) { $email } else { '(no email attribute)' }
            "$($id.PSChildName -replace '^.*_','') -> $emailMasked"
        }

        New-CheckResult -Category 'Authentication' -Check 'Microsoft 365 Identity' -Status 'PASS' `
            -Value "$($identities.Count) identity/identities found: $($entries -join '; ')" `
            -Expected 'At least one signed-in Microsoft 365 identity' `
            -Finding 'Microsoft 365 identity is registered with Office on this machine.' `
            -Command $cmd
    }
}

function Get-WebAccountManagerStatus {
    Invoke-SafeCheck -Category 'Authentication' -Check 'Web Account Manager (WAM)' -ScriptBlock {
        $cmd = "Get-Service TokenBroker"
        $svc = Get-Service -Name 'TokenBroker' -ErrorAction SilentlyContinue

        if (-not $svc) {
            return New-CheckResult -Category 'Authentication' -Check 'Web Account Manager (WAM)' -Status 'WARNING' `
                -Value 'TokenBroker service not found' -Expected 'TokenBroker (WAM) service present and running/auto' `
                -Finding 'The Web Account Manager (TokenBroker) service, used for modern Microsoft 365 authentication, could not be found.' `
                -Command $cmd
        }

        $status = if ($svc.Status -eq 'Running' -or $svc.StartType -eq 'Manual') { 'PASS' } else { 'WARNING' }
        $rec = if ($status -eq 'WARNING') { 'TokenBroker is not running and is not set to start on demand. Modern (OAuth/WAM-based) sign-in to Microsoft 365 apps may fail. Consider setting the service to Manual/Automatic start.' } else { '' }

        New-CheckResult -Category 'Authentication' -Check 'Web Account Manager (WAM)' -Status $status `
            -Value "Status: $($svc.Status) | StartType: $($svc.StartType)" `
            -Expected 'Running or available on demand (Manual start type is normal)' `
            -Finding "$(if($status -eq 'PASS'){'WAM/TokenBroker service is available for modern authentication.'}else{'WAM/TokenBroker service may not be available when needed.'})" `
            -Recommendation $rec -Command $cmd
    }
}

function Get-CredentialManagerSummary {
    Invoke-SafeCheck -Category 'Authentication' -Check 'Credential Manager Entries' -ScriptBlock {
        $cmd = 'cmdkey /list'
        $raw = & cmdkey.exe /list 2>&1 | Out-String

        # Extract only target names, never the credential blob itself.
        $targets = [regex]::Matches($raw, 'Target:\s*(\S+)') | ForEach-Object { $_.Groups[1].Value }
        $officeRelated = $targets | Where-Object { $_ -match 'Office|Outlook|MicrosoftAccount|AzureAD|MSOfficeSSO|WebAccount' }

        if ($officeRelated.Count -eq 0) {
            return New-CheckResult -Category 'Authentication' -Check 'Credential Manager Entries' -Status 'INFO' `
                -Value 'No Office/Outlook-related credential entries found' -Expected 'Informational' `
                -Finding 'No cached Office/Microsoft 365 credentials found in Windows Credential Manager for this user. This can be normal for token-based (WAM) sign-in.' `
                -Command $cmd
        }

        New-CheckResult -Category 'Authentication' -Check 'Credential Manager Entries' -Status 'INFO' `
            -Value "$($officeRelated.Count) Office/Microsoft-related entr$(if($officeRelated.Count -eq 1){'y'}else{'ies'}) found" `
            -Expected 'Informational' `
            -Finding "Cached credential entries detected: $($officeRelated -join '; '). No credential values were read or modified. Stale or duplicate entries here occasionally cause repeated sign-in prompts." `
            -Recommendation 'If the user reports repeated/looping sign-in prompts, stale entries here (Control Panel > Credential Manager) are a common manual remediation step - not performed automatically by this tool.' `
            -Command $cmd
    }
}

function Get-OfficeRegistryIdentityKeys {
    Invoke-SafeCheck -Category 'Authentication' -Check 'Office Identity Registry' -ScriptBlock {
        $cmd = 'HKCU:\Software\Microsoft\Office\16.0\Common\SignIn'
        $signInPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\SignIn'
        $signInInfo = Get-RegistryValueSafe -Path $signInPath -Name 'SignInOptions'

        New-CheckResult -Category 'Authentication' -Check 'Office Identity Registry' -Status 'INFO' `
            -Value "SignInOptions: $(if($null -ne $signInInfo){$signInInfo}else{'Not set (default)'})" `
            -Expected 'Informational' `
            -Finding 'Office sign-in configuration collected for context.' -Command $cmd
    }
}

function Invoke-AllIdentityChecks {
    @(
        Get-M365IdentityInfo
        Get-WebAccountManagerStatus
        Get-CredentialManagerSummary
        Get-OfficeRegistryIdentityKeys
    )
}
