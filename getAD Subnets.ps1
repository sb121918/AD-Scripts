# --- FUNCTIONS ---
function Get-ForestSiteDCs {
    param (
        [string]$TargetSiteName,
        [array]$GlobalDCList
    )
    # Filter the pre-cached global list for the specific site
    # This handles both the simple name and the DistinguishedName format
    $SiteDCs = $GlobalDCList | Where-Object { 
        $_.Site -eq $TargetSiteName -or 
        $_.Site -like "CN=$TargetSiteName,CN=Sites,*" 
    }
    return $SiteDCs
}

# --- MAIN SCRIPT ---

# 1. Get ALL DCs from the entire forest root (Global Catalog) once to maximize speed
Write-Host "Gathering every DC from the entire forest (Root: $((Get-ADForest).RootDomain))..." -ForegroundColor Cyan
try {
    $GlobalDCList = Get-ADDomainController -Filter * -Server (Get-ADForest).RootDomain
} catch {
    Write-Error "Failed to connect to Forest Root. Ensure you have connectivity to the Root Domain."
    return
}

# 2. Get all AD Subnets
Write-Host "Querying AD Subnets..." -ForegroundColor Cyan
$Subnets = Get-ADReplicationSubnet -Filter * -Properties Site, Description, Location

# 3. Process the Report
$Report = ForEach ($Subnet in $Subnets) {
    # Extract the Site Name from the Subnet property
    $SiteName = if ($Subnet.Site) { 
        ($Subnet.Site -split ",")[0].Replace("CN=", "") 
    } else { 
        "Unassigned" 
    }

    $DcNames = "N/A"
    $Count = 0

    if ($SiteName -ne "Unassigned") {
        # Call the internal function
        $SiteDCs = Get-ForestSiteDCs -TargetSiteName $SiteName -GlobalDCList $GlobalDCList
        $Count = ($SiteDCs | Measure-Object).Count
        $DcNames = ($SiteDCs.Name | Sort-Object) -join ", "
    }

    [PSCustomObject]@{
        "Subnet Name"        = $Subnet.Name
        "AD Site"            = $SiteName
        "DC Count"           = $Count
        "DC Names"           = $DcNames
        "Subnet Description" = $Subnet.Description
        "Subnet Location"    = $Subnet.Location
    }
}

# 4. Final Output
$Report | Sort-Object "AD Site" | Out-GridView -Title "Forest-Wide Subnet & DC Audit (Combined)"

Write-Host "Report Complete." -ForegroundColor Green