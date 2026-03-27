param(
    [int]$Limit = 0,
    [string]$TargetDomain
)

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

    $serverParam = @{}
    if ($TargetDomain) { $serverParam['Server'] = $TargetDomain }

    # Get all site links with all properties
    try {
        $siteLinks = Get-ADReplicationSiteLink -Filter * -Properties * @serverParam -ErrorAction Stop
    } catch {
        $siteLinks = Get-ADReplicationSiteLink -Filter * @serverParam -ErrorAction Stop
    }

    $Results = foreach ($link in $siteLinks) {
        $dns = @($link.SiteList)

        $siteNames = $dns |
          ForEach-Object {
            if ($_ -and ($_ -match 'CN=([^,]+)')) { $matches[1] } else { $_ }
          } |
          Where-Object { $_ } |
          Sort-Object

        # Get replication interval - try multiple property names
        $replicationInterval = $null
        if ($link.PSObject.Properties.Name -contains 'replInterval') {
            $replicationInterval = [int]$link.replInterval
        } elseif ($link.PSObject.Properties.Name -contains 'Interval') {
            $replicationInterval = [int]$link.Interval
        } elseif ($link.PSObject.Properties.Name -contains 'ReplicationInterval') {
            if ($link.ReplicationInterval -is [TimeSpan]) { 
                $replicationInterval = [int]$link.ReplicationInterval.TotalMinutes 
            } else { 
                $replicationInterval = [int]$link.ReplicationInterval 
            }
        }

        # Get replication schedule info
        $scheduleInfo = "24x7"
        if ($link.PSObject.Properties.Name -contains 'Schedule' -and $link.Schedule) {
            $scheduleInfo = "Custom Schedule"
        }

        # Get options (change notification, etc.)
        $options = $link.Options
        $changeNotification = $false
        if ($options -band 1) {
            $changeNotification = $true
        }

        [PSCustomObject]@{
            SiteLinkName           = $link.Name
            Description            = $link.Description
            Cost                   = $link.Cost
            ReplicationIntervalMin = $replicationInterval
            ReplicationSchedule    = $scheduleInfo
            ChangeNotification     = $changeNotification
            SiteCount              = $siteNames.Count
            AssociatedSites        = $siteNames -join ', '
            DistinguishedName      = $link.DistinguishedName
            Created                = Format-DateForJSON $link.Created
            Modified               = Format-DateForJSON $link.Modified
        }
    }

    @($Results) | ConvertTo-Json -Depth 5 -Compress
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
