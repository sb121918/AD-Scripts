param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    if ($TargetDomain) {
        $forest = Get-ADForest -Server $TargetDomain
    } else {
        $forest = Get-ADForest
    }
    $results = @()

    foreach ($domainName in $forest.Domains) {
        try {
            $domain = Get-ADDomain -Identity $domainName -ErrorAction SilentlyContinue
            
            # Count DCs
            $dcCount = 0
            try {
                $dcs = Get-ADDomainController -Filter * -Server $domainName -ErrorAction SilentlyContinue
                $dcCount = @($dcs).Count
            } catch {}

            $results += [PSCustomObject]@{
                DomainName = $domain.DNSRoot
                NetBIOSName = $domain.NetBIOSName
                DomainMode = $domain.DomainMode.ToString()
                ParentDomain = if ($domain.ParentDomain) { $domain.ParentDomain } else { $null }
                ChildDomains = if ($domain.ChildDomains) { ($domain.ChildDomains -join ", ") } else { $null }
                DCCount = $dcCount
                DistinguishedName = $domain.DistinguishedName
                DomainSID = if ($domain.DomainSID) { $domain.DomainSID.ToString() } else { $null }
                InfrastructureMaster = $domain.InfrastructureMaster
                RIDMaster = $domain.RIDMaster
                PDCEmulator = $domain.PDCEmulator
            }
        } catch {
            $results += [PSCustomObject]@{
                DomainName = $domainName
                NetBIOSName = "Error"
                DomainMode = "Unable to query"
                ParentDomain = $null
                ChildDomains = $null
                DCCount = 0
                DistinguishedName = $null
                DomainSID = $null
                InfrastructureMaster = $null
                RIDMaster = $null
                PDCEmulator = $null
            }
        }
    }

    $results | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
