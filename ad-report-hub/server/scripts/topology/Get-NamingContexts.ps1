param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    if ($TargetDomain) {
        $forest = Get-ADForest -Server $TargetDomain
        $rootDSE = Get-ADRootDSE -Server $TargetDomain
    } else {
        $forest = Get-ADForest
        $rootDSE = Get-ADRootDSE
    }

    $results = @()

    # Schema NC
    $results += [PSCustomObject]@{
        Name = "Schema"
        Type = "Schema Partition"
        DistinguishedName = $rootDSE.schemaNamingContext
    }

    # Configuration NC
    $results += [PSCustomObject]@{
        Name = "Configuration"
        Type = "Configuration Partition"
        DistinguishedName = $rootDSE.configurationNamingContext
    }

    # Domain NCs
    foreach ($domainName in $forest.Domains) {
        try {
            $domain = Get-ADDomain -Identity $domainName -ErrorAction SilentlyContinue
            $results += [PSCustomObject]@{
                Name = $domain.DNSRoot
                Type = "Domain Partition"
                DistinguishedName = $domain.DistinguishedName
            }
        } catch {}
    }

    # Application Partitions
    foreach ($appPartition in $forest.ApplicationPartitions) {
        $partitionName = $appPartition -replace "DC=", "" -replace ",", "."
        $results += [PSCustomObject]@{
            Name = $partitionName
            Type = "Application Partition"
            DistinguishedName = $appPartition
        }
    }

    $results | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
