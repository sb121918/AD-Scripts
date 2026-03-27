param([string]$TargetDomain)

$ErrorActionPreference = 'SilentlyContinue'

function Format-DateForJSON {
    param($Date)
    if ($null -eq $Date -or $Date -eq [DateTime]::MinValue) { return $null }
    try {
        if ($Date -is [DateTime]) {
            return $Date.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        }
        $parsed = [DateTime]::Parse($Date)
        return $parsed.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    } catch { return $null }
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$serverParam = @{}
if ($TargetDomain) { $serverParam['Server'] = $TargetDomain }

$AllTrusts = New-Object System.Collections.Generic.List[PSObject]

try {
    $Forest = Get-ADForest @serverParam -ErrorAction Stop
    $RootDomain = $Forest.RootDomain
    $AllDomains = if ($TargetDomain) { @($TargetDomain) } else { $Forest.Domains }

    # Get Forest-level trusts from the root domain
    try {
        $ForestTrusts = Get-ADTrust -Filter * -Server $RootDomain -ErrorAction SilentlyContinue
        
        foreach ($Trust in $ForestTrusts) {
            $TrustObject = [PSCustomObject]@{
                SourceDomain      = $RootDomain
                SourceLevel       = "Forest Root"
                TargetDomain      = $Trust.Target
                TrustType         = $Trust.TrustType
                TrustDirection    = switch ($Trust.Direction) {
                    0 { "Disabled" }
                    1 { "Inbound" }
                    2 { "Outbound" }
                    3 { "Bidirectional" }
                    default { $Trust.Direction }
                }
                TrustAttributes   = $Trust.TrustAttributes
                ForestTransitive  = $Trust.ForestTransitive
                SelectiveAuth     = $Trust.SelectiveAuthentication
                SIDFilteringForestAware = $Trust.SIDFilteringForestAware
                SIDFilteringQuarantined = $Trust.SIDFilteringQuarantined
                DisallowTransivity = $Trust.DisallowTransivity
                IntraForest       = $Trust.IntraForest
                IsTreeParent      = $Trust.IsTreeParent
                IsTreeRoot        = $Trust.IsTreeRoot
                TGTDelegation     = $Trust.TGTDelegation
                UplevelOnly       = $Trust.UplevelOnly
                UsesAESKeys       = $Trust.UsesAESKeys
                UsesRC4Encryption = $Trust.UsesRC4Encryption
                Created           = Format-DateForJSON $Trust.Created
                Modified          = Format-DateForJSON $Trust.Modified
                DistinguishedName = $Trust.DistinguishedName
            }
            $AllTrusts.Add($TrustObject)
        }
    } catch {
        # Continue to check child domains
    }

    # Get trusts from each child domain
    foreach ($Domain in $AllDomains) {
        if ($Domain -eq $RootDomain) { continue }  # Already processed root domain
        
        try {
            $DomainTrusts = Get-ADTrust -Filter * -Server $Domain -ErrorAction SilentlyContinue
            
            foreach ($Trust in $DomainTrusts) {
                # Check if this trust already exists (avoid duplicates for intra-forest trusts)
                $Exists = $AllTrusts | Where-Object { 
                    $_.SourceDomain -eq $Domain -and $_.TargetDomain -eq $Trust.Target 
                }
                
                if (-not $Exists) {
                    $TrustObject = [PSCustomObject]@{
                        SourceDomain      = $Domain
                        SourceLevel       = "Child Domain"
                        TargetDomain      = $Trust.Target
                        TrustType         = $Trust.TrustType
                        TrustDirection    = switch ($Trust.Direction) {
                            0 { "Disabled" }
                            1 { "Inbound" }
                            2 { "Outbound" }
                            3 { "Bidirectional" }
                            default { $Trust.Direction }
                        }
                        TrustAttributes   = $Trust.TrustAttributes
                        ForestTransitive  = $Trust.ForestTransitive
                        SelectiveAuth     = $Trust.SelectiveAuthentication
                        SIDFilteringForestAware = $Trust.SIDFilteringForestAware
                        SIDFilteringQuarantined = $Trust.SIDFilteringQuarantined
                        DisallowTransivity = $Trust.DisallowTransivity
                        IntraForest       = $Trust.IntraForest
                        IsTreeParent      = $Trust.IsTreeParent
                        IsTreeRoot        = $Trust.IsTreeRoot
                        TGTDelegation     = $Trust.TGTDelegation
                        UplevelOnly       = $Trust.UplevelOnly
                        UsesAESKeys       = $Trust.UsesAESKeys
                        UsesRC4Encryption = $Trust.UsesRC4Encryption
                        Created           = Format-DateForJSON $Trust.Created
                        Modified          = Format-DateForJSON $Trust.Modified
                        DistinguishedName = $Trust.DistinguishedName
                    }
                    $AllTrusts.Add($TrustObject)
                }
            }
        } catch {
            # Continue with next domain
        }
    }
}
catch {
    @{ Error = "Failed to query forest trusts: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

if ($AllTrusts.Count -eq 0) {
    @() | ConvertTo-Json
} else {
    @($AllTrusts) | ConvertTo-Json -Depth 5
}
