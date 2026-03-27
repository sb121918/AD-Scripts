param(
    [int]$Limit = 0,
    [string]$TargetDomain
)

# Helper function to format dates as ISO 8601 strings for JavaScript
function Format-DateForJSON {
    param($Date)
    if ($null -eq $Date) { return $null }
    try {
        return $Date.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    } catch {
        return $null
    }
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $serverParam = @{}
    if ($TargetDomain) { $serverParam['Server'] = $TargetDomain }
    
    $ConfigPartition = (Get-ADRootDSE @serverParam).configurationNamingContext

    # Pre-fetch Site Links and DCs for mapping
    $SiteLinks = Get-ADReplicationSiteLink -Filter * @serverParam
    $SiteServerObjects = Get-ADObject -SearchBase "CN=Sites,$ConfigPartition" -Filter "objectClass -eq 'server'" @serverParam

    # Get Subnets with Extended Properties (no limit - fetch all)
    $Subnets = Get-ADReplicationSubnet -Filter * -Properties Site, Description, Location, WhenCreated, WhenChanged @serverParam

    $Results = foreach ($Subnet in $Subnets) {
        $SiteDN = $Subnet.Site
        $SiteName = if ($SiteDN) { ($SiteDN -split ",")[0].Replace("CN=", "") } else { "Unassigned" }

        # Find the Site Link associated with this Site
        $AssociatedLink = $SiteLinks | Where-Object { $_.SiteList -match "CN=$SiteName,CN=Sites" }
        
        # Count DCs using the Configuration Partition logic
        $Count = 0
        if ($SiteName -ne "Unassigned") {
            $Count = ($SiteServerObjects | Where-Object { $_.DistinguishedName -like "*CN=$SiteName,CN=Sites,$ConfigPartition" }).Count
        }

        [PSCustomObject]@{
            Subnet           = $Subnet.Name
            ADSite           = $SiteName
            DCCount          = $Count
            SiteLink         = $AssociatedLink.Name
            LinkCost         = $AssociatedLink.Cost
            Description      = $Subnet.Description
            PhysicalLocation = $Subnet.Location
            CreatedDate      = Format-DateForJSON $Subnet.WhenCreated
            LastModified     = Format-DateForJSON $Subnet.WhenChanged
        }
    }

    @($Results) | ConvertTo-Json -Depth 5 -Compress
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
