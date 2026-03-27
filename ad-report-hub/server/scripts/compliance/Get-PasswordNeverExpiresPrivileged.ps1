# Get-PasswordNeverExpiresPrivileged.ps1
# Finds accounts in privileged groups with "Password Never Expires" flag
param(
    [string]$ForestDomain = "",
    [string]$TargetDomain = "",
    [string]$PrivilegedGroupsCsv = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    # Avoid warning noise corrupting JSON output
    $WarningPreference = 'SilentlyContinue'

    $forest = $null
    try {
        if ($ForestDomain) {
            $forest = Get-ADForest -Server $ForestDomain
        } else {
            $forest = Get-ADForest
        }
    } catch {
        $forest = $null
    }
    $results = @()

    $privilegedGroups = if ($PrivilegedGroupsCsv) {
        $PrivilegedGroupsCsv.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    } else {
        @(
            'Domain Admins',
            'Enterprise Admins',
            'Schema Admins',
            'Administrators',
            'Account Operators',
            'Backup Operators'
        )
    }

    $domainsToQuery = if ($TargetDomain) { @($TargetDomain) } elseif ($forest -and $forest.Domains) { $forest.Domains } elseif ($ForestDomain) { @($ForestDomain) } else { @() }

    foreach ($domain in $domainsToQuery) {
        try {
            $dc = $null
            try {
                $domainInfo = Get-ADDomain -Server $domain -ErrorAction Stop
                $dc = $domainInfo.PDCEmulator
            } catch {
                $dc = (Get-ADDomainController -DomainName $domain -Discover -ErrorAction Stop).HostName
            }

            foreach ($groupName in $privilegedGroups) {
                try {
                    $members = Get-ADGroupMember -Identity $groupName -Server $dc -Recursive -ErrorAction SilentlyContinue |
                        Where-Object { $_.objectClass -eq 'user' }

                    foreach ($member in $members) {
                        try {
                            $user = Get-ADUser -Identity $member.SamAccountName -Server $dc -Properties PasswordNeverExpires, Enabled, PasswordLastSet, DisplayName, LastLogonDate -ErrorAction SilentlyContinue

                            if ($user.PasswordNeverExpires -eq $true) {
                                $passwordAge = if ($user.PasswordLastSet) { [math]::Round(((Get-Date) - $user.PasswordLastSet).TotalDays) } else { 'Unknown' }

                                $results += [PSCustomObject]@{
                                    Name                  = $user.DisplayName
                                    SamAccountName        = $user.SamAccountName
                                    PrivilegedGroup       = $groupName
                                    Domain                = $domain
                                    Enabled               = $user.Enabled
                                    PasswordNeverExpires   = $true
                                    PasswordLastSet       = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Never' }
                                    PasswordAgeDays       = $passwordAge
                                    LastLogonDate         = if ($user.LastLogonDate) { $user.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
                                    Risk                  = if ($passwordAge -is [int] -and $passwordAge -gt 365) { 'High' } elseif ($passwordAge -is [int] -and $passwordAge -gt 90) { 'Medium' } else { 'Low' }
                                }
                            }
                        } catch {
                            # Skip individual user errors
                        }
                    }
                } catch {
                    # Skip group errors
                }
            }
        } catch {
            # Skip domain errors
        }
    }

    $results = $results | Sort-Object SamAccountName, PrivilegedGroup -Unique

    if ($results.Count -eq 0) {
        $results = @()
    }

    $results | ConvertTo-Json -Depth 5
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
}
