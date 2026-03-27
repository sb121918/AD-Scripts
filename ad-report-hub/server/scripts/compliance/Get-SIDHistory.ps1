# Get-SIDHistory.ps1
# Finds accounts with SID history attributes (potential security risk from migrations)
param(
    [string]$ForestDomain = "",
    [string]$TargetDomain = "",
    [int]$MaxResults = 5000
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

    $domainsToQuery = if ($TargetDomain) { @($TargetDomain) } elseif ($forest -and $forest.Domains) { $forest.Domains } elseif ($ForestDomain) { @($ForestDomain) } else { @() }

    foreach ($domain in $domainsToQuery) {
        try {
            # Prefer querying the PDC emulator for consistent results
            $dc = $null
            try {
                $domainInfo = Get-ADDomain -Server $domain -ErrorAction Stop
                $dc = $domainInfo.PDCEmulator
            } catch {
                $dc = (Get-ADDomainController -DomainName $domain -Discover -ErrorAction Stop).HostName
            }

            # Find users with SID history (use LDAP filter for performance)
            $usersWithSIDHistory = Get-ADUser -LDAPFilter '(sIDHistory=*)' -Server $dc `
              -ResultPageSize 2000 -ResultSetSize $MaxResults `
              -Properties SIDHistory, DisplayName, SamAccountName, Enabled, LastLogonDate, WhenCreated -ErrorAction SilentlyContinue

            foreach ($user in $usersWithSIDHistory) {
                $sidHistoryCount = @($user.SIDHistory).Count
                $sidHistoryValues = ($user.SIDHistory | ForEach-Object { $_.Value }) -join '; '

                $results += [PSCustomObject]@{
                    Name             = $user.DisplayName
                    SamAccountName   = $user.SamAccountName
                    ObjectType       = 'User'
                    Domain           = $domain
                    Enabled          = $user.Enabled
                    SIDHistoryCount  = $sidHistoryCount
                    SIDHistory       = $sidHistoryValues
                    LastLogonDate    = if ($user.LastLogonDate) { $user.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
                    WhenCreated      = $user.WhenCreated.ToString('yyyy-MM-dd')
                    Risk             = if ($sidHistoryCount -gt 3) { 'High' } elseif ($sidHistoryCount -gt 1) { 'Medium' } else { 'Low' }
                }
            }

            # Find groups with SID history
            $groupsWithSIDHistory = Get-ADGroup -LDAPFilter '(sIDHistory=*)' -Server $dc `
              -ResultPageSize 2000 -ResultSetSize $MaxResults `
              -Properties SIDHistory, Name, SamAccountName, WhenCreated, GroupScope, GroupCategory -ErrorAction SilentlyContinue

            foreach ($group in $groupsWithSIDHistory) {
                $sidHistoryCount = @($group.SIDHistory).Count
                $sidHistoryValues = ($group.SIDHistory | ForEach-Object { $_.Value }) -join '; '

                $results += [PSCustomObject]@{
                    Name             = $group.Name
                    SamAccountName   = $group.SamAccountName
                    ObjectType       = 'Group'
                    Domain           = $domain
                    Enabled          = "$($group.GroupScope) / $($group.GroupCategory)"
                    SIDHistoryCount  = $sidHistoryCount
                    SIDHistory       = $sidHistoryValues
                    LastLogonDate    = '-'
                    WhenCreated      = $group.WhenCreated.ToString('yyyy-MM-dd')
                    Risk             = if ($sidHistoryCount -gt 3) { 'High' } elseif ($sidHistoryCount -gt 1) { 'Medium' } else { 'Low' }
                }
            }

            # Find computers with SID history
            $computersWithSIDHistory = Get-ADComputer -LDAPFilter '(sIDHistory=*)' -Server $dc `
              -ResultPageSize 2000 -ResultSetSize $MaxResults `
              -Properties SIDHistory, Name, SamAccountName, Enabled, WhenCreated -ErrorAction SilentlyContinue

            foreach ($comp in $computersWithSIDHistory) {
                $sidHistoryCount = @($comp.SIDHistory).Count
                $sidHistoryValues = ($comp.SIDHistory | ForEach-Object { $_.Value }) -join '; '

                $results += [PSCustomObject]@{
                    Name             = $comp.Name
                    SamAccountName   = $comp.SamAccountName
                    ObjectType       = 'Computer'
                    Domain           = $domain
                    Enabled          = $comp.Enabled
                    SIDHistoryCount  = $sidHistoryCount
                    SIDHistory       = $sidHistoryValues
                    LastLogonDate    = '-'
                    WhenCreated      = $comp.WhenCreated.ToString('yyyy-MM-dd')
                    Risk             = if ($sidHistoryCount -gt 3) { 'High' } elseif ($sidHistoryCount -gt 1) { 'Medium' } else { 'Low' }
                }
            }
        } catch {
            # Skip domain errors
        }
    }

    if ($results.Count -eq 0) {
        $results = @()
    }

    $results | ConvertTo-Json -Depth 5
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
}
