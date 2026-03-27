param(
    [int]$Days = 30,
    [string]$ForestDomain = "",
    [string]$TargetDomain = "",
    [string]$PrivilegedGroupsCsv = "",
    [int]$MaxEventsPerDomain = 2000,
    [int]$MaxEventsPerDC = 250
)

$ErrorActionPreference = 'SilentlyContinue'

function Get-EventDataMap {
    param([xml]$EventXml)
    $map = @{}
    foreach ($d in $EventXml.Event.EventData.Data) {
        try {
            $name = [string]$d.Name
            if ($name) { $map[$name] = [string]$d.'#text' }
        } catch { }
    }
    return $map
}

function Get-CnFromDnOrSidLike {
    param([string]$Value)
    if (-not $Value) { return $null }
    # Often "CN=User Name,OU=...,DC=..." - pull the CN for readability
    if ($Value -match '^CN=([^,]+),') { return $Matches[1] }
    return $Value
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$cutoff = (Get-Date).AddDays(-$Days)

$forest = $null
try {
    if ($ForestDomain) { $forest = Get-ADForest -Server $ForestDomain -ErrorAction Stop }
    else { $forest = Get-ADForest -ErrorAction Stop }
} catch { $forest = $null }

$privilegedGroups = if ($PrivilegedGroupsCsv) {
    $PrivilegedGroupsCsv.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
} else {
    @(
        'Domain Admins',
        'Enterprise Admins',
        'Schema Admins',
        'Administrators',
        'Account Operators',
        'Backup Operators',
        'Server Operators',
        'Print Operators'
    )
}

$domainsToQuery = if ($TargetDomain) { @($TargetDomain) } elseif ($forest -and $forest.Domains) { @($forest.Domains) } elseif ($ForestDomain) { @($ForestDomain) } else { @() }

# Relevant Security log events for group membership changes
# 4728/4729: member added/removed from security-enabled global group
# 4732/4733: member added/removed from security-enabled local group
# 4756/4757: member added/removed from security-enabled universal group
$eventIds = @(4728,4729,4732,4733,4756,4757)
$addedIds = @(4728,4732,4756)
$removedIds = @(4729,4733,4757)

$results = @()
$errors = @()

foreach ($domain in $domainsToQuery) {
    try {
        $pdc = $null
        try {
            $domainInfo = Get-ADDomain -Server $domain -ErrorAction Stop
            $pdc = $domainInfo.PDCEmulator
        } catch {
            $pdc = (Get-ADDomainController -DomainName $domain -Discover -ErrorAction Stop).HostName
        }

        # IMPORTANT: these events are logged on the DC that processed the change.
        # Querying *all* DCs can be very slow. Instead, use replication metadata to identify the originating DCs
        # for privileged group membership changes within the window, then query Security logs on only those DCs.
        $dcs = @()
        $originDcs = New-Object System.Collections.Generic.HashSet[string]

        foreach ($groupName in $privilegedGroups) {
            try {
                $group = Get-ADGroup -Identity $groupName -Server $pdc -ErrorAction SilentlyContinue
                if (-not $group) { continue }

                $meta = $null
                try {
                    $meta = Get-ADReplicationAttributeMetadata -Object $group.DistinguishedName -Server $pdc -Properties member -ErrorAction Stop |
                        Where-Object { $_.AttributeName -eq 'member' } |
                        Select-Object -First 1
                } catch { $meta = $null }

                if ($meta -and $meta.LastOriginatingChangeTime -ge $cutoff) {
                    $origin = [string]$meta.LastOriginatingChangeDirectoryServerIdentity
                    $originHost = $null
                    # Common case: already a hostname (e.g. vaddc01.domain.com)
                    if ($origin -match '^[A-Za-z0-9_.-]+$') {
                        $originHost = $origin
                    } elseif ($origin -match 'CN=NTDS Settings,CN=([^,]+),') {
                        $originHost = $Matches[1]
                    }

                    if ($originHost) { [void]$originDcs.Add($originHost) }
                }
            } catch { }
        }

        if ($originDcs.Count -gt 0) {
            $dcs = @($originDcs)
        } else {
            # Fallback: query PDC only (fastest)
            $dcs = @($pdc)
        }

        $domainEventCount = 0

        foreach ($dc in ($dcs | Where-Object { $_ })) {
            if ($domainEventCount -ge $MaxEventsPerDomain) { break }

            $events = $null
            try {
                $events = Get-WinEvent -ComputerName $dc -FilterHashtable @{
                    LogName   = 'Security'
                    Id        = $eventIds
                    StartTime = $cutoff
                } -MaxEvents $MaxEventsPerDC -ErrorAction Stop
            } catch {
                $errors += [PSCustomObject]@{
                    Domain = $domain
                    DomainController = $dc
                    Status = 'ERROR'
                    Error = $_.Exception.Message
                }
                continue
            }

            foreach ($evt in ($events | Where-Object { $_ })) {
                if ($domainEventCount -ge $MaxEventsPerDomain) { break }
                try {
                    $xml = [xml]$evt.ToXml()
                    $data = Get-EventDataMap -EventXml $xml

                    $targetGroup = $data['TargetUserName']
                    if (-not $targetGroup) { continue }

                    # Filter to privileged groups only (case-insensitive)
                    $isPriv = $false
                    foreach ($pg in $privilegedGroups) {
                        if ($targetGroup -ieq $pg) { $isPriv = $true; break }
                    }
                    if (-not $isPriv) { continue }

                    $memberRaw = $data['MemberName']
                    $member = Get-CnFromDnOrSidLike -Value $memberRaw

                    $who = $null
                    $subjUser = $data['SubjectUserName']
                    $subjDom = $data['SubjectDomainName']
                    if ($subjUser) {
                        if ($subjDom) { $who = "$subjDom\\$subjUser" } else { $who = $subjUser }
                    }

                    $action = if ($addedIds -contains $evt.Id) { 'Added' } elseif ($removedIds -contains $evt.Id) { 'Removed' } else { 'Changed' }

                    $results += [PSCustomObject]@{
                        GroupName     = $targetGroup
                        Domain        = $domain
                        Action        = $action
                        Member        = $member
                        MemberRaw     = $memberRaw
                        ChangedBy     = $who
                        ChangedBySid  = $data['SubjectUserSid']
                        ChangeTime    = $evt.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                        EventId       = $evt.Id
                        DomainController = $dc
                        Status        = 'OK'
                    }
                    $domainEventCount++
                } catch { }
            }
        }
    } catch {
        $errors += [PSCustomObject]@{
            Domain = $domain
            DomainController = $null
            Status = 'ERROR'
            Error = $_.Exception.Message
        }
    }
}

# If we got no events, return error diagnostics too (so UI shows why it's empty).
if (($results.Count -eq 0) -and ($errors.Count -gt 0)) {
    $errors | ConvertTo-Json -Depth 5 -Compress
    exit 0
}

$results | Sort-Object ChangeTime -Descending | ConvertTo-Json -Depth 5 -Compress

