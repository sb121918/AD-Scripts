# Get-KerberosDelegation.ps1
# Finds accounts with Kerberos delegation settings (unconstrained, constrained, RBCD)
param(
    [string]$ForestDomain = "",
    [string]$TargetDomain = ""
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

            # Find accounts with unconstrained delegation (TrustedForDelegation)
            $unconstrained = Get-ADObject -Filter { TrustedForDelegation -eq $true } -Server $dc -Properties Name, ObjectClass, SamAccountName, TrustedForDelegation, Enabled, WhenCreated, Description -ErrorAction SilentlyContinue

            foreach ($obj in $unconstrained) {
                # Skip domain controllers (they have unconstrained delegation by default)
                if ($obj.ObjectClass -eq 'computer') {
                    $isDC = Get-ADDomainController -Filter "Name -eq '$($obj.Name)'" -Server $dc -ErrorAction SilentlyContinue
                    if ($isDC) { continue }
                }

                $results += [PSCustomObject]@{
                    Name             = $obj.Name
                    SamAccountName   = $obj.SamAccountName
                    ObjectType       = $obj.ObjectClass
                    DelegationType   = 'Unconstrained'
                    Domain           = $domain
                    Enabled          = $obj.Enabled
                    DelegationTarget = 'Any Service'
                    Description      = $obj.Description
                    WhenCreated      = if ($obj.WhenCreated) { $obj.WhenCreated.ToString('yyyy-MM-dd') } else { '-' }
                    Risk             = 'Critical'
                }
            }

            # Find accounts with constrained delegation
            $constrained = Get-ADObject -Filter { msDS-AllowedToDelegateTo -like '*' } -Server $dc -Properties Name, ObjectClass, SamAccountName, 'msDS-AllowedToDelegateTo', Enabled, WhenCreated, Description -ErrorAction SilentlyContinue

            foreach ($obj in $constrained) {
                $targets = ($obj.'msDS-AllowedToDelegateTo') -join '; '
                $results += [PSCustomObject]@{
                    Name             = $obj.Name
                    SamAccountName   = $obj.SamAccountName
                    ObjectType       = $obj.ObjectClass
                    DelegationType   = 'Constrained'
                    Domain           = $domain
                    Enabled          = $obj.Enabled
                    DelegationTarget = $targets
                    Description      = $obj.Description
                    WhenCreated      = if ($obj.WhenCreated) { $obj.WhenCreated.ToString('yyyy-MM-dd') } else { '-' }
                    Risk             = 'Medium'
                }
            }

            # Find accounts with Resource-Based Constrained Delegation
            $rbcd = Get-ADObject -Filter { msDS-AllowedToActOnBehalfOfOtherIdentity -like '*' } -Server $dc -Properties Name, ObjectClass, SamAccountName, Enabled, WhenCreated, Description -ErrorAction SilentlyContinue

            foreach ($obj in $rbcd) {
                $results += [PSCustomObject]@{
                    Name             = $obj.Name
                    SamAccountName   = $obj.SamAccountName
                    ObjectType       = $obj.ObjectClass
                    DelegationType   = 'Resource-Based (RBCD)'
                    Domain           = $domain
                    Enabled          = $obj.Enabled
                    DelegationTarget = 'See msDS-AllowedToActOnBehalfOfOtherIdentity'
                    Description      = $obj.Description
                    WhenCreated      = if ($obj.WhenCreated) { $obj.WhenCreated.ToString('yyyy-MM-dd') } else { '-' }
                    Risk             = 'Medium'
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
