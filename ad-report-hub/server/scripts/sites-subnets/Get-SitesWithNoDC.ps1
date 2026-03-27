param([string]$TargetDomain)

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
    
    # Get all sites
    $AllSites = Get-ADReplicationSite -Filter * -Properties WhenCreated, WhenChanged, Description @serverParam
    
    # Get all server objects (DCs) and their sites
    $SiteServerObjects = Get-ADObject -SearchBase "CN=Sites,$ConfigPartition" -Filter "objectClass -eq 'server'" -Properties DistinguishedName @serverParam
    
    # Get site links for additional info
    $SiteLinks = Get-ADReplicationSiteLink -Filter * @serverParam
    
    $Results = @()
    
    foreach ($Site in $AllSites) {
        $SiteName = $Site.Name
        
        # Count DCs in this site
        $DCCount = ($SiteServerObjects | Where-Object { $_.DistinguishedName -like "*CN=$SiteName,CN=Sites,$ConfigPartition" }).Count
        
        # Only include sites with NO domain controllers
        if ($DCCount -eq 0) {
            # Find associated site links
            $AssociatedLinks = $SiteLinks | Where-Object { $_.SiteList -match "CN=$SiteName,CN=Sites" }
            $LinkNames = ($AssociatedLinks | Select-Object -ExpandProperty Name) -join ", "
            
            # Count subnets associated with this site
            $SubnetCount = (Get-ADReplicationSubnet -Filter "Site -eq '$($Site.DistinguishedName)'" @serverParam -ErrorAction SilentlyContinue | Measure-Object).Count
            
            $Results += [PSCustomObject]@{
                SiteName        = $SiteName
                Description     = $Site.Description
                DCCount         = 0
                SubnetCount     = $SubnetCount
                SiteLinks       = if ($LinkNames) { $LinkNames } else { "None" }
                Created         = Format-DateForJSON $Site.WhenCreated
                LastModified    = Format-DateForJSON $Site.WhenChanged
                DistinguishedName = $Site.DistinguishedName
            }
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 5 -Compress
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
