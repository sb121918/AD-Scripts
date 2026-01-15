# 1. Identify Configuration Context
$ConfigPartition = (Get-ADRootDSE).configurationNamingContext

Write-Host "Gathering deep metadata for all Forest Subnets..." -ForegroundColor Cyan

# 2. Pre-fetch Site Links and DCs for mapping
$SiteLinks = Get-ADReplicationSiteLink -Filter *
$SiteServerObjects = Get-ADObject -SearchBase "CN=Sites,$ConfigPartition" -Filter "objectClass -eq 'server'"

# 3. Get Subnets with Extended Properties
$Subnets = Get-ADReplicationSubnet -Filter * -Properties Site, Description, Location, WhenCreated, WhenChanged

$Report = ForEach ($Subnet in $Subnets) {
    # Extract Site Name
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
        "Subnet"           = $Subnet.Name
        "AD Site"          = $SiteName
        "DC Count"         = $Count
        #"Site Link"        = $AssociatedLink.Name
        #"Link Cost"        = $AssociatedLink.Cost
        "Description"      = $Subnet.Description
        "Physical Location"= $Subnet.Location
        "Created Date"     = $Subnet.WhenCreated
        "Last Modified"    = $Subnet.WhenChanged
    }
}

$Report | Sort-Object "AD Site" | Out-GridView -Title "Advanced AD Subnet Lifecycle Audit"