# Get-AdminSDHolderProtected.ps1
# Finds accounts protected by AdminSDHolder (adminCount=1)
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

            # Find users with adminCount=1
            $protectedUsers = Get-ADUser -Filter { adminCount -eq 1 } -Server $dc -Properties adminCount, DisplayName, SamAccountName, Enabled, LastLogonDate, PasswordLastSet, MemberOf, WhenCreated -ErrorAction SilentlyContinue

            foreach ($user in $protectedUsers) {
                # Check if user is actually in a privileged group
                $privilegedMemberships = @()
                $knownPrivileged = if ($PrivilegedGroupsCsv) {
                    $PrivilegedGroupsCsv.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                } else {
                    @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators')
                }

                $matchEnterpriseSystems = ($knownPrivileged | Where-Object { $_ -match 'Enterprise\\s*Systems\\s*Admin' }).Count -gt 0

                foreach ($memberOf in $user.MemberOf) {
                    $groupName = ($memberOf -split ',')[0] -replace 'CN=', ''
                    if ($knownPrivileged -contains $groupName) {
                        $privilegedMemberships += $groupName
                    } elseif ($matchEnterpriseSystems -and ($groupName -match 'Enterprise\\s*Systems\\s*Admin')) {
                        # Handle naming variations for this custom group
                        $privilegedMemberships += $groupName
                    }
                }

                $isOrphaned = $privilegedMemberships.Count -eq 0
                $passwordAge = if ($user.PasswordLastSet) { [math]::Round(((Get-Date) - $user.PasswordLastSet).TotalDays) } else { 'Unknown' }

                $results += [PSCustomObject]@{
                    Name                    = $user.DisplayName
                    SamAccountName          = $user.SamAccountName
                    Domain                  = $domain
                    Enabled                 = $user.Enabled
                    AdminCount              = $user.adminCount
                    IsOrphanedAdminCount    = $isOrphaned
                    PrivilegedGroups        = if ($privilegedMemberships.Count -gt 0) { $privilegedMemberships -join '; ' } else { 'None (orphaned)' }
                    LastLogonDate           = if ($user.LastLogonDate) { $user.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
                    PasswordLastSet         = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Never' }
                    PasswordAgeDays         = $passwordAge
                    WhenCreated             = $user.WhenCreated.ToString('yyyy-MM-dd')
                    Risk                    = if ($isOrphaned) { 'High' } elseif (-not $user.Enabled) { 'Medium' } else { 'Low' }
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
