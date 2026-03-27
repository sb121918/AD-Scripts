param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $WarningPreference = 'SilentlyContinue'
    $ErrorActionPreference = 'SilentlyContinue'

    function Get-SiteName {
        param($Value)
        if ($null -eq $Value) { return $null }
        $s = [string]$Value
        if (-not $s) { return $null }
        if ($s -match 'CN=([^,]+)') { return $matches[1] }
        return $s
    }

    # Anchor to the selected forest/domain
    $forest = $null
    try {
        if ($TargetDomain) { $forest = Get-ADForest -Server $TargetDomain -ErrorAction Stop }
        else { $forest = Get-ADForest -ErrorAction Stop }
    } catch { $forest = $null }

    $serverParam = @{}
    if ($TargetDomain) { $serverParam['Server'] = $TargetDomain }

    $domainsToQuery = @()
    if ($forest -and $forest.Domains) {
        $domainsToQuery = @($forest.Domains)
    } elseif ($TargetDomain) {
        $domainsToQuery = @($TargetDomain)
    }

    # Sites (forest-wide, config partition)
    $sites = Get-ADReplicationSite -Filter * -Properties Description,WhenCreated,WhenChanged @serverParam

    # Subnets (forest-wide, config partition)
    $subnets = Get-ADReplicationSubnet -Filter * -Properties Site @serverParam
    $subnetCountsBySite = @{}
    foreach ($sn in @($subnets)) {
        $siteName = Get-SiteName $sn.Site
        if (-not $siteName) { continue }
        if (-not $subnetCountsBySite.ContainsKey($siteName)) { $subnetCountsBySite[$siteName] = 0 }
        $subnetCountsBySite[$siteName]++
    }

    # DCs (need to loop domains; DC Site property is already the site name)
    $dcCountsBySite = @{}
    foreach ($d in $domainsToQuery) {
        try {
            $dcs = Get-ADDomainController -Filter * -Server $d -ErrorAction SilentlyContinue
            foreach ($dc in @($dcs)) {
                $siteName = [string]$dc.Site
                if (-not $siteName) { continue }
                if (-not $dcCountsBySite.ContainsKey($siteName)) { $dcCountsBySite[$siteName] = 0 }
                $dcCountsBySite[$siteName]++
            }
        } catch { }
    }

    # Site links (forest-wide, config partition)
    $siteLinks = Get-ADReplicationSiteLink -Filter * -Properties * @serverParam
    $siteLinkCountsBySite = @{}
    $linkedSitesBySite = @{} # site -> HashSet(otherSites)

    foreach ($link in @($siteLinks)) {
        $siteNames = @($link.SiteList) |
            ForEach-Object { Get-SiteName $_ } |
            Where-Object { $_ } |
            Sort-Object -Unique

        if ($siteNames.Count -lt 1) { continue }

        foreach ($s in $siteNames) {
            if (-not $siteLinkCountsBySite.ContainsKey($s)) { $siteLinkCountsBySite[$s] = 0 }
            $siteLinkCountsBySite[$s]++

            if (-not $linkedSitesBySite.ContainsKey($s)) {
                $linkedSitesBySite[$s] = New-Object System.Collections.Generic.HashSet[string]
            }

            foreach ($other in $siteNames) {
                if ($other -and $other -ne $s) { [void]$linkedSitesBySite[$s].Add([string]$other) }
            }
        }
    }

    $results = foreach ($site in @($sites)) {
        $name = [string]$site.Name
        $linked = @()
        if ($linkedSitesBySite.ContainsKey($name)) { $linked = @($linkedSitesBySite[$name]) }

        [PSCustomObject]@{
            SiteName          = $name
            DCCount           = if ($dcCountsBySite.ContainsKey($name)) { [int]$dcCountsBySite[$name] } else { 0 }
            SubnetCount       = if ($subnetCountsBySite.ContainsKey($name)) { [int]$subnetCountsBySite[$name] } else { 0 }
            SiteLinksCount    = if ($siteLinkCountsBySite.ContainsKey($name)) { [int]$siteLinkCountsBySite[$name] } else { 0 }
            LinkedSitesCount  = $linked.Count
            LinkedSites       = ($linked | Sort-Object) -join ', '
            Description       = $site.Description
            WhenCreated       = if ($site.WhenCreated) { $site.WhenCreated.ToString('yyyy-MM-dd') } else { $null }
            WhenChanged       = if ($site.WhenChanged) { $site.WhenChanged.ToString('yyyy-MM-dd') } else { $null }
        }
    }

    @($results) | Sort-Object SiteName | ConvertTo-Json -Depth 6 -Compress
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}

