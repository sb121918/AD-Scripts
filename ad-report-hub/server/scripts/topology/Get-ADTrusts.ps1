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

    # Get trusts for each domain
    foreach ($domainName in $forest.Domains) {
        try {
            $trusts = Get-ADTrust -Filter * -Server $domainName -ErrorAction SilentlyContinue

            foreach ($trust in $trusts) {
                $direction = switch ($trust.Direction) {
                    0 { "Disabled" }
                    1 { "Inbound" }
                    2 { "Outbound" }
                    3 { "Bidirectional" }
                    default { $trust.Direction.ToString() }
                }

                $trustType = switch ($trust.TrustType) {
                    1 { "Downlevel" }
                    2 { "Uplevel" }
                    3 { "MIT" }
                    4 { "DCE" }
                    default { $trust.TrustType.ToString() }
                }

                $attributes = @()
                if ($trust.ForestTransitive) { $attributes += "Forest" }
                if ($trust.SelectiveAuthentication) { $attributes += "Selective Auth" }
                if ($trust.SIDFilteringQuarantined) { $attributes += "SID Filtered" }
                if ($trust.SIDFilteringForestAware) { $attributes += "Forest Aware" }

                $results += [PSCustomObject]@{
                    SourceDomain = $domainName
                    TargetDomain = $trust.Target
                    TrustType = $trustType
                    TrustDirection = $direction
                    TrustAttributes = if ($attributes.Count -gt 0) { $attributes -join ", " } else { "Standard" }
                    IsTransitive = $trust.ForestTransitive -or ($trust.TrustType -eq 2)
                    WhenCreated = if ($trust.Created) { $trust.Created.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
                    DistinguishedName = $trust.DistinguishedName
                }
            }
        } catch {
            # Skip domains we can't query
        }
    }

    if ($results.Count -eq 0) {
        $results = @([PSCustomObject]@{
            SourceDomain = $forest.RootDomain
            TargetDomain = "No trusts found"
            TrustType = "-"
            TrustDirection = "-"
            TrustAttributes = "-"
            IsTransitive = $false
            WhenCreated = $null
            DistinguishedName = $null
        })
    }

    $results | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
