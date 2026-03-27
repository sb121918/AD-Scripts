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

    $results = @(
        [PSCustomObject]@{ Property = "Forest Name"; Value = $forest.Name }
        [PSCustomObject]@{ Property = "Forest Mode"; Value = $forest.ForestMode.ToString() }
        [PSCustomObject]@{ Property = "Root Domain"; Value = $forest.RootDomain }
        [PSCustomObject]@{ Property = "Schema Master"; Value = $forest.SchemaMaster }
        [PSCustomObject]@{ Property = "Domain Naming Master"; Value = $forest.DomainNamingMaster }
        [PSCustomObject]@{ Property = "Global Catalog Servers"; Value = ($forest.GlobalCatalogs -join ", ") }
        [PSCustomObject]@{ Property = "Total Domains"; Value = $forest.Domains.Count }
        [PSCustomObject]@{ Property = "Domains"; Value = ($forest.Domains -join ", ") }
        [PSCustomObject]@{ Property = "Sites"; Value = ($forest.Sites -join ", ") }
        [PSCustomObject]@{ Property = "Application Partitions"; Value = ($forest.ApplicationPartitions -join ", ") }
        [PSCustomObject]@{ Property = "UPN Suffixes"; Value = if ($forest.UPNSuffixes) { ($forest.UPNSuffixes -join ", ") } else { "None" } }
        [PSCustomObject]@{ Property = "SPN Suffixes"; Value = if ($forest.SPNSuffixes) { ($forest.SPNSuffixes -join ", ") } else { "None" } }
    )

    $results | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
