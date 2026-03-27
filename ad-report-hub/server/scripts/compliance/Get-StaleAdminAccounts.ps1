# Get-StaleAdminAccounts.ps1
# Finds privileged users who haven't logged on in X days
param(
    [int]$Days = 90,
    [string]$ForestDomain = "",
    [string]$TargetDomain = "",
    [string]$PrivilegedGroupsCsv = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    # Avoid warning noise corrupting JSON output
    $WarningPreference = 'SilentlyContinue'

    $cutoffDate = (Get-Date).AddDays(-$Days)
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
            'Administrators'
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
                            $user = Get-ADUser -Identity $member.SamAccountName -Server $dc -Properties LastLogonDate, Enabled, PasswordLastSet, WhenCreated, DisplayName -ErrorAction SilentlyContinue

                            $isStale = $false
                            if (-not $user.LastLogonDate) {
                                $isStale = $true
                            } elseif ($user.LastLogonDate -lt $cutoffDate) {
                                $isStale = $true
                            }

                            if ($isStale) {
                                $daysSinceLogon = if ($user.LastLogonDate) { [math]::Round(((Get-Date) - $user.LastLogonDate).TotalDays) } else { 'Never' }

                                $results += [PSCustomObject]@{
                                    Name             = $user.DisplayName
                                    SamAccountName   = $user.SamAccountName
                                    PrivilegedGroup  = $groupName
                                    Domain           = $domain
                                    Enabled          = $user.Enabled
                                    LastLogonDate    = if ($user.LastLogonDate) { $user.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
                                    DaysSinceLogon   = $daysSinceLogon
                                    PasswordLastSet  = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Never' }
                                    WhenCreated      = $user.WhenCreated.ToString('yyyy-MM-dd')
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

    # Deduplicate by SamAccountName (user may be in multiple groups)
    $results = $results | Sort-Object SamAccountName, PrivilegedGroup -Unique

    if ($results.Count -eq 0) {
        $results = @()
    }

    $results | ConvertTo-Json -Depth 5
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
}
