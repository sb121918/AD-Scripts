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
    
    # Get all subnets where Site is null or empty (not associated with any site)
    $Subnets = Get-ADReplicationSubnet -Filter * -Properties Site, Description, Location, WhenCreated, WhenChanged @serverParam
    
    $Results = @()
    
    foreach ($Subnet in $Subnets) {
        # Check if subnet has no site assigned
        if ($null -eq $Subnet.Site -or $Subnet.Site -eq "") {
            $Results += [PSCustomObject]@{
                Subnet           = $Subnet.Name
                ADSite           = "Not Assigned"
                Description      = $Subnet.Description
                PhysicalLocation = $Subnet.Location
                Created          = Format-DateForJSON $Subnet.WhenCreated
                LastModified     = Format-DateForJSON $Subnet.WhenChanged
                DistinguishedName = $Subnet.DistinguishedName
            }
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 5 -Compress
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
