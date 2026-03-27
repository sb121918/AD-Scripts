param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $results = @()

    # Get all domain controllers
    if ($TargetDomain) {
        $forest = Get-ADForest -Server $TargetDomain
    } else {
        $forest = Get-ADForest
    }
    
    foreach ($domainName in $forest.Domains) {
        try {
            $dcs = Get-ADDomainController -Filter * -Server $domainName -ErrorAction SilentlyContinue
            
            foreach ($dc in $dcs) {
                try {
                    # Get replication partners and status
                    $replPartners = Get-ADReplicationPartnerMetadata -Target $dc.HostName -ErrorAction SilentlyContinue
                    
                    if ($replPartners) {
                        foreach ($partner in $replPartners) {
                            # Determine status
                            $status = "Success"
                            $failureCount = 0
                            
                            if ($partner.LastReplicationResult -ne 0) {
                                $status = "Failed"
                                $failureCount = $partner.ConsecutiveReplicationFailures
                            }
                            
                            # Format last replication time
                            $lastRepl = "Never"
                            if ($partner.LastReplicationSuccess) {
                                $lastRepl = $partner.LastReplicationSuccess.ToString("yyyy-MM-dd HH:mm:ss")
                            }

                            # Extract source DC name
                            $sourceDC = $partner.Partner
                            if ($sourceDC -match 'CN=NTDS Settings,CN=([^,]+)') {
                                $sourceDC = $matches[1]
                            }

                            # Extract naming context name
                            $ncName = $partner.Partition
                            if ($ncName -match 'DC=([^,]+)') {
                                $ncName = $matches[1]
                            } elseif ($ncName -match 'CN=([^,]+)') {
                                $ncName = $matches[1]
                            }

                            $results += [PSCustomObject]@{
                                SourceDC          = $sourceDC
                                DestinationDC     = $dc.HostName
                                Domain            = $domainName
                                NamingContext     = $ncName
                                LastReplication   = $lastRepl
                                ReplicationStatus = $status
                                FailureCount      = $failureCount
                                LastAttempt       = if ($partner.LastReplicationAttempt) { $partner.LastReplicationAttempt.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                                ResultCode        = $partner.LastReplicationResult
                            }
                        }
                    }
                } catch {
                    # If we can't get replication info for this DC, add a placeholder
                    $results += [PSCustomObject]@{
                        SourceDC          = "Unable to query"
                        DestinationDC     = $dc.HostName
                        Domain            = $domainName
                        NamingContext     = "N/A"
                        LastReplication   = "N/A"
                        ReplicationStatus = "Unknown"
                        FailureCount      = 0
                        LastAttempt       = "N/A"
                        ResultCode        = -1
                    }
                }
            }
        } catch {
            # Skip domains we can't query
        }
    }

    if ($results.Count -eq 0) {
        $results = @([PSCustomObject]@{
            SourceDC          = "No replication data found"
            DestinationDC     = "-"
            Domain            = "-"
            NamingContext     = "-"
            LastReplication   = "-"
            ReplicationStatus = "Unknown"
            FailureCount      = 0
            LastAttempt       = "-"
            ResultCode        = 0
        })
    }

    $results | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
