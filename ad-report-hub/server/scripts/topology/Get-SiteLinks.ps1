param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $siteLinksParams = @{
        Filter     = '*'
        Properties = '*'
        ErrorAction = 'Stop'
    }
    if ($TargetDomain) {
        $siteLinksParams['Server'] = $TargetDomain
    }
    $siteLinks = Get-ADReplicationSiteLink @siteLinksParams

    $results = foreach ($link in $siteLinks) {
        $sites = @($link.SiteList) | ForEach-Object {
            if ($_ -match 'CN=([^,]+)') { $matches[1] } else { $_ }
        } | Sort-Object

        # Get replication interval
        $interval = 180
        if ($link.PSObject.Properties.Name -contains 'replInterval') {
            $interval = [int]$link.replInterval
        } elseif ($link.PSObject.Properties.Name -contains 'ReplicationInterval') {
            $interval = [int]$link.ReplicationInterval
        }

        # Check for change notification
        $changeNotify = $false
        if ($link.Options -band 1) {
            $changeNotify = $true
        }

        # Schedule info
        $schedule = "24x7"
        if ($link.Schedule) {
            $schedule = "Custom Schedule"
        }

        [PSCustomObject]@{
            SiteLinkName        = $link.Name
            Sites               = $sites -join ", "
            SiteCount           = $sites.Count
            Cost                = $link.Cost
            ReplicationInterval = $interval
            Schedule            = $schedule
            ChangeNotification  = $changeNotify
            Description         = $link.Description
            DistinguishedName   = $link.DistinguishedName
        }
    }

    @($results) | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
