<#
.SYNOPSIS
    Gets the SID of every domain in the AD forest.

.DESCRIPTION
    Enumerates all domains in the current (or specified) forest and returns
    each domain's name and Domain SID (S-1-5-21-...). Useful for mapping
    domain names to SIDs or validating which domain a SID belongs to.

.PARAMETER TargetDomain
    Optional. DNS name of any domain in the forest (e.g. "contoso.com").
    If omitted, uses the current user's domain to discover the forest.

.EXAMPLE
    .\Get-ForestDomainSIDs.ps1
.EXAMPLE
    .\Get-ForestDomainSIDs.ps1 -TargetDomain "contoso.com"
#>

param(
    [string]$TargetDomain = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $ErrorActionPreference = 'Stop'

    if ($TargetDomain) {
        $forest = Get-ADForest -Server $TargetDomain
    } else {
        $forest = Get-ADForest
    }

    $results = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($domainName in $forest.Domains) {
        try {
            $domain = Get-ADDomain -Identity $domainName -ErrorAction Stop
            $sidStr = if ($domain.DomainSID) { $domain.DomainSID.ToString() } else { $null }
            $results.Add([PSCustomObject]@{
                DomainName  = $domain.DNSRoot
                NetBIOSName = $domain.NetBIOSName
                DomainSID   = $sidStr
            })
        } catch {
            $results.Add([PSCustomObject]@{
                DomainName  = $domainName
                NetBIOSName = $null
                DomainSID   = $null
                Error       = $_.Exception.Message
            })
        }
    }

    $results | ConvertTo-Json -Depth 3
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
