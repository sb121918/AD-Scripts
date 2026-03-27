param(
    [int]$Limit = 0,
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $serverParam = @{}
    if ($TargetDomain) { $serverParam['Server'] = $TargetDomain }
    
    # Get ALL DCs from the entire forest
    $rootDomain = if ($TargetDomain) { $TargetDomain } else { (Get-ADForest @serverParam).RootDomain }
    $GlobalDCList = Get-ADDomainController -Filter * -Server $rootDomain

    # Get all AD Subnets (no limit - fetch all)
    $Subnets = Get-ADReplicationSubnet -Filter * -Properties Site, Description, Location @serverParam

    $Results = foreach ($Subnet in $Subnets) {
        $SiteName = if ($Subnet.Site) { 
            ($Subnet.Site -split ",")[0].Replace("CN=", "") 
        } else { 
            "Unassigned" 
        }

        $DcNames = "N/A"
        $Count = 0

        if ($SiteName -ne "Unassigned") {
            $SiteDCs = $GlobalDCList | Where-Object { 
                $_.Site -eq $SiteName -or 
                $_.Site -like "CN=$SiteName,CN=Sites,*" 
            }
            $Count = ($SiteDCs | Measure-Object).Count
            $DcNames = ($SiteDCs.Name | Sort-Object) -join ", "
        }

        [PSCustomObject]@{
            SubnetName  = $Subnet.Name
            ADSite      = $SiteName
            DCCount     = $Count
            DCNames     = $DcNames
            Description = $Subnet.Description
            Location    = $Subnet.Location
        }
    }

    @($Results) | ConvertTo-Json -Depth 5 -Compress
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
