try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $results = @()

    # Get all replication site links first
    $siteLinks = Get-ADReplicationSiteLink -Filter * -Properties * -ErrorAction SilentlyContinue

    foreach ($link in $siteLinks) {
        $sites = @($link.SiteList) | ForEach-Object {
            if ($_ -match 'CN=([^,]+)') { $matches[1] } else { $_ }
        }

        # Get replication interval
        $interval = 180 # Default
        if ($link.PSObject.Properties.Name -contains 'replInterval') {
            $interval = $link.replInterval
        } elseif ($link.PSObject.Properties.Name -contains 'ReplicationInterval') {
            $interval = $link.ReplicationInterval
        }

        # For each pair of sites, create a connection entry
        for ($i = 0; $i -lt $sites.Count; $i++) {
            for ($j = $i + 1; $j -lt $sites.Count; $j++) {
                $results += [PSCustomObject]@{
                    SourceServer = $sites[$i]
                    DestinationServer = $sites[$j]
                    NamingContext = "All Partitions"
                    ReplicationType = "Intersite"
                    Schedule = "$interval min"
                    SiteLinkName = $link.Name
                    Cost = $link.Cost
                }
            }
        }
    }

    # Also try to get intrasite connections
    try {
        $connections = Get-ADReplicationConnection -Filter * -ErrorAction SilentlyContinue
        
        foreach ($conn in $connections) {
            $sourceName = $conn.ReplicateFromDirectoryServer -replace "CN=NTDS Settings,CN=", "" -replace ",CN=.*", ""
            $destName = $conn.ReplicateToDirectoryServer -replace "CN=NTDS Settings,CN=", "" -replace ",CN=.*", ""

            $results += [PSCustomObject]@{
                SourceServer = $sourceName
                DestinationServer = $destName
                NamingContext = "All Partitions"
                ReplicationType = "Intrasite"
                Schedule = "Immediate"
                SiteLinkName = "N/A"
                Cost = 0
            }
        }
    } catch {
        # Intrasite connections might not be available
    }

    if ($results.Count -eq 0) {
        $results = @([PSCustomObject]@{
            SourceServer = "No connections found"
            DestinationServer = "-"
            NamingContext = "-"
            ReplicationType = "-"
            Schedule = "-"
            SiteLinkName = "-"
            Cost = 0
        })
    }

    $results | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
