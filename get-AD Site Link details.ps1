Import-Module ActiveDirectory -ErrorAction Stop

Write-Host "Querying AD Site Links from the Configuration Partition..." -ForegroundColor Cyan

try {
    $siteLinks = Get-ADReplicationSiteLink -Filter * -Properties Description, Cost, Interval, SiteList -ErrorAction Stop
} catch {
    Write-Warning "Couldn't request 'Interval' property; retrying without it."
    $siteLinks = Get-ADReplicationSiteLink -Filter * -Properties Description, Cost, SiteList -ErrorAction Stop
}

$report = foreach ($link in $siteLinks) {
    $dns = @($link.SiteList)

    # Fixed: use ForEach-Object pipeline (avoids the empty-pipe parser error)
    $siteNames = $dns |
      ForEach-Object {
        if ($_ -and ($_ -match 'CN=([^,]+)')) { $matches[1] } else { $_ }
      } |
      Where-Object { $_ } |
      Sort-Object

    if ($link.PSObject.Properties.Name -contains 'Interval') {
        $durationMin = [int]$link.Interval
    } elseif ($link.PSObject.Properties.Name -contains 'ReplicationInterval') {
        if ($link.ReplicationInterval -is [TimeSpan]) { $durationMin = [int]$link.ReplicationInterval.TotalMinutes }
        else { $durationMin = [int]$link.ReplicationInterval }
    } else {
        $durationMin = $null
    }

    [PSCustomObject]@{
        'Site Link Name'    = $link.Name
        'Description'       = $link.Description
        'Cost'              = $link.Cost
        'Duration (Min)'    = $durationMin
        'Site Count'        = $siteNames.Count
        'Associated Sites'  = $siteNames -join ', '
    }
}

if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
    $report | Sort-Object 'Site Link Name' | Out-GridView -Title 'Active Directory Site Link Topology Audit'
} else {
    $report | Sort-Object 'Site Link Name' | Format-Table -AutoSize
}