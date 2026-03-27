# Get-PrivilegedGroupChanges.ps1
# Queries privileged group membership changes via replication metadata
param(
    [int]$Days = 30,
    [string]$ForestDomain = "",
    [string]$TargetDomain = "",
    [string]$PrivilegedGroupsCsv = "",
    [switch]$IncludeMemberDetails
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    function Get-CnFromMemberDn {
        param([string]$Dn)
        if ([string]::IsNullOrWhiteSpace($Dn)) { return $null }
        if ($Dn -match '^CN=([^,]+),') { return $Matches[1] }
        return $Dn
    }

    function Normalize-ReplicationChangeTimeUtc {
        param($T)
        if ($null -eq $T) { return $null }
        try {
            $d = [datetime]$T
            if ($d.Kind -eq [DateTimeKind]::Unspecified) {
                $d = [datetime]::SpecifyKind($d, [DateTimeKind]::Local)
            }
            return $d.ToUniversalTime()
        } catch {
            return $null
        }
    }

    function Get-MemberDnFromLinkMetaRow {
        param($Entry)
        if (-not $Entry) { return $null }
        foreach ($name in @('AttributeValue', 'Value', 'Object')) {
            if ($Entry.PSObject.Properties.Name -contains $name) {
                $v = $Entry.$name
                if ($null -eq $v) { continue }
                if ($v -is [string]) {
                    $ts = $v.Trim()
                    if ($ts.Length -gt 0) { return $ts }
                }
            }
        }
        foreach ($p in $Entry.PSObject.Properties) {
            $n = $p.Name
            if ($n -match '^(AttributeName|Version|LastOriginating|IsDeleted|IsLink|ObjectType|LocalChange|OriginatingUsn|uSNChanged|AttributeSyntax)') { continue }
            $v = $p.Value
            if ($v -is [string] -and $v -match '(?i)^CN=.+,') { return $v.Trim() }
        }
        return $null
    }

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

    # Fallback when linked-value replication metadata is not available: same event IDs as Get-PrivilegedGroupMembershipAudit.ps1
    function Get-MemberRowsFromSecurityLog {
        param(
            [string]$ResolvedGroupName,
            [string]$DomainDns,
            [string]$PdcHostname,
            [string]$OriginatingDcIdentity,
            [datetime]$StartTime,
            [int]$MaxPerDc = 400
        )

        $eventIds = @(4728, 4729, 4732, 4733, 4756, 4757)
        $addedIds = @(4728, 4732, 4756)
        $removedIds = @(4729, 4733, 4757)

        $dcCandidates = @()
        if ($PdcHostname) { $dcCandidates += $PdcHostname.Trim() }

        $origin = [string]$OriginatingDcIdentity
        if ($origin -match '^[A-Za-z0-9_.-]+$') {
            $dcCandidates += $origin.Trim()
        } elseif ($origin -match 'CN=NTDS Settings,CN=([^,]+),') {
            $dcCandidates += $Matches[1].Trim()
        }

        $dcs = @($dcCandidates | Where-Object { $_ } | Select-Object -Unique)

        $domainNetbios = $null
        try {
            $domainNetbios = (Get-ADDomain -Server $DomainDns -ErrorAction Stop).NetBIOSName
        } catch { }

        $rows = New-Object System.Collections.Generic.List[object]

        foreach ($targetDc in $dcs) {
            $events = $null
            try {
                $events = @(Get-WinEvent -ComputerName $targetDc -FilterHashtable @{
                    LogName   = 'Security'
                    Id        = $eventIds
                    StartTime = $StartTime
                } -MaxEvents $MaxPerDc -ErrorAction Stop)
            } catch {
                continue
            }

            foreach ($evt in $events) {
                try {
                    $xml = [xml]$evt.ToXml()
                    $data = Get-EventDataMap -EventXml $xml

                    $targetGroup = $data['TargetUserName']
                    if (-not $targetGroup -or ($targetGroup -ine $ResolvedGroupName)) { continue }

                    if ($domainNetbios -and $data['TargetDomainName'] -and ($data['TargetDomainName'] -ine $domainNetbios)) {
                        continue
                    }

                    $memberRaw = $data['MemberName']
                    if ([string]::IsNullOrWhiteSpace($memberRaw)) { continue }

                    $memberCn = Get-CnFromMemberDn -Dn $memberRaw
                    if ([string]::IsNullOrWhiteSpace($memberCn)) {
                        $memberCn = $memberRaw
                    }

                    $action = if ($addedIds -contains $evt.Id) {
                        'Added'
                    } elseif ($removedIds -contains $evt.Id) {
                        'Removed'
                    } else {
                        'Changed'
                    }

                    $rows.Add([PSCustomObject]@{
                        GroupName           = $ResolvedGroupName
                        Domain              = $DomainDns
                        AttributeChanged    = 'member (Security event)'
                        ChangeTime          = $evt.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                        OriginatingDC       = $targetDc
                        ChangeVersion       = $evt.Id
                        ChangeType          = $action
                        MemberDN            = $memberRaw
                        MemberCN            = $memberCn
                    })
                } catch { }
            }
        }

        return @($rows | Sort-Object { try { [datetime]$_.ChangeTime } catch { [datetime]'1900-01-01' } } -Descending)
    }

    $cutoffDate = (Get-Date).AddDays(-$Days)
    $cutoffUtc = $cutoffDate.ToUniversalTime()
    $forest = $null
    try {
        if ($ForestDomain) {
            $forest = Get-ADForest -Server $ForestDomain
        } else {
            $forest = Get-ADForest
        }
    } catch {
        $forest = $null
    }
    $results = @()

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

    $domainsToQuery = if ($TargetDomain) { @($TargetDomain) } elseif ($forest -and $forest.Domains) { $forest.Domains } elseif ($ForestDomain) { @($ForestDomain) } else { @() }

    foreach ($domain in $domainsToQuery) {
        try {
            # Prefer querying the PDC emulator for consistent results
            $dc = $null
            try {
                $domainInfo = Get-ADDomain -Server $domain -ErrorAction Stop
                $dc = $domainInfo.PDCEmulator
            } catch {
                $dc = (Get-ADDomainController -DomainName $domain -Discover -ErrorAction Stop).HostName
            }

            foreach ($groupName in $privilegedGroups) {
                try {
                    # Use -Identity (faster/safer than filter) and fetch WhenChanged for fallback mode
                    $resolvedGroupName = $groupName
                    $group = $null

                    if ($groupName -eq 'AD-Enterprise Systems Admins') {
                        # This custom group has naming variations in the environment; search robustly (Viacom-only)
                        $candidateNames = @(
                            "AD-Enterprise Systems Admins",
                            "AD-Enterprise-Systems Admins",
                            "AD-Enterprise-System Admin",
                            "AD-Enterprise System Admins",
                            "Enterprise Systems Admins"
                        )

                        foreach ($n in $candidateNames) {
                            try {
                                $group = Get-ADGroup -Filter "Name -eq '$n'" -Server $dc -Properties WhenChanged -ErrorAction SilentlyContinue
                                if ($group) { $resolvedGroupName = $group.Name; break }
                            } catch { }
                        }

                        if (-not $group) {
                            try {
                                $group = Get-ADGroup -Filter "Name -like '*Enterprise*System*Admin*'" -Server $dc -Properties WhenChanged -ErrorAction SilentlyContinue | Select-Object -First 1
                                if ($group) { $resolvedGroupName = $group.Name }
                            } catch { }
                        }
                    } else {
                        $group = Get-ADGroup -Identity $groupName -Server $dc -Properties WhenChanged -ErrorAction SilentlyContinue
                    }
                    if (-not $group) { continue }

                    $wroteAny = $false
                    $memberAttrMeta = $null
                    try {
                        # Preferred (accurate): replication metadata for the 'member' attribute (detects add/remove occurred).
                        $memberAttrMeta = Get-ADReplicationAttributeMetadata `
                            -Object $group.DistinguishedName `
                            -Server $dc `
                            -Properties member `
                            -ErrorAction Stop |
                            Where-Object { $_.AttributeName -eq 'member' } |
                            Select-Object -First 1
                    } catch {
                        $memberAttrMeta = $null
                    }

                    if ($memberAttrMeta -and ($memberAttrMeta.LastOriginatingChangeTime -ge $cutoffDate)) {
                        if ($IncludeMemberDetails) {
                            # Per-member adds/removes via linked-value replication metadata (PDC).
                            $replMetadata = @()
                            try {
                                $replMetadata = @(
                                    Get-ADReplicationAttributeMetadata `
                                        -Object $group.DistinguishedName `
                                        -Server $dc `
                                        -Properties member `
                                        -ShowAllLinkedValues `
                                        -IncludeDeletedObjects `
                                        -ErrorAction Stop |
                                        Where-Object { $_.AttributeName -ieq 'member' }
                                )
                            } catch {
                                try {
                                    $replMetadata = @(
                                        Get-ADReplicationAttributeMetadata `
                                            -Object $group.DistinguishedName `
                                            -Server $dc `
                                            -Properties member `
                                            -ShowAllLinkedValues `
                                            -ErrorAction Stop |
                                            Where-Object { $_.AttributeName -ieq 'member' }
                                    )
                                } catch {
                                    $replMetadata = @()
                                }
                            }

                            $detailRows = New-Object System.Collections.Generic.List[object]
                            foreach ($entry in $replMetadata) {
                                $t = $entry.LastOriginatingChangeTime
                                $inWindow = ($t -ge $cutoffDate)
                                if (-not $inWindow) {
                                    $entryUtc = Normalize-ReplicationChangeTimeUtc $t
                                    $inWindow = ($null -ne $entryUtc -and $entryUtc -ge $cutoffUtc)
                                }
                                if (-not $inWindow) { continue }

                                $memberValue = Get-MemberDnFromLinkMetaRow $entry
                                if ([string]::IsNullOrWhiteSpace($memberValue)) { continue }

                                $isDeleted = $false
                                if ($entry.PSObject.Properties.Name -contains 'IsDeleted') {
                                    $isDeleted = [bool]$entry.IsDeleted
                                }

                                $changeType = if ($isDeleted) { 'Removed' } else { 'Added' }

                                $detailRows.Add([PSCustomObject]@{
                                    GroupName           = $resolvedGroupName
                                    Domain              = $domain
                                    AttributeChanged    = 'member'
                                    ChangeTime          = $entry.LastOriginatingChangeTime.ToString('yyyy-MM-dd HH:mm:ss')
                                    OriginatingDC       = $entry.LastOriginatingChangeDirectoryServerIdentity
                                    ChangeVersion       = $entry.Version
                                    ChangeType          = $changeType
                                    MemberDN            = $memberValue
                                    MemberCN            = (Get-CnFromMemberDn $memberValue)
                                })
                            }

                            if ($detailRows.Count -gt 0) {
                                foreach ($r in $detailRows) { $results += $r }
                            } else {
                                # Linked-value metadata unavailable (common). Resolve members from DC Security events 4728/4729/4732/4733/4756/4757.
                                $secRows = @(Get-MemberRowsFromSecurityLog `
                                    -ResolvedGroupName $resolvedGroupName `
                                    -DomainDns $domain `
                                    -PdcHostname $dc `
                                    -OriginatingDcIdentity $memberAttrMeta.LastOriginatingChangeDirectoryServerIdentity `
                                    -StartTime $cutoffDate `
                                    -MaxPerDc 400)

                                if ($secRows.Count -gt 0) {
                                    foreach ($sr in $secRows) { $results += $sr }
                                } else {
                                    $results += [PSCustomObject]@{
                                        GroupName        = $resolvedGroupName
                                        Domain           = $domain
                                        AttributeChanged = $memberAttrMeta.AttributeName
                                        ChangeTime       = $memberAttrMeta.LastOriginatingChangeTime.ToString('yyyy-MM-dd HH:mm:ss')
                                        OriginatingDC    = $memberAttrMeta.LastOriginatingChangeDirectoryServerIdentity
                                        ChangeVersion    = $memberAttrMeta.Version
                                        ChangeType       = '—'
                                        MemberDN         = $null
                                        MemberCN         = $null
                                        DetailNote       = 'Membership changed in this period, but neither linked replication values nor matching Security events (read Security log on the PDC and originating DC) returned the member. Confirm auditing of group membership changes, the lookback window, and that this account can read Security logs on those DCs; or run Privileged Group Membership Audit.'
                                    }
                                }
                            }
                        } else {
                            # Default (fast): one row per group when membership changed.
                            $results += [PSCustomObject]@{
                                GroupName           = $resolvedGroupName
                                Domain              = $domain
                                AttributeChanged    = $memberAttrMeta.AttributeName
                                ChangeTime          = $memberAttrMeta.LastOriginatingChangeTime.ToString('yyyy-MM-dd HH:mm:ss')
                                OriginatingDC       = $memberAttrMeta.LastOriginatingChangeDirectoryServerIdentity
                                ChangeVersion       = $memberAttrMeta.Version
                            }
                        }
                        $wroteAny = $true
                    }

                    # Fallback: if replication metadata isn't available (common permissions issue),
                    # use WhenChanged as a best-effort indicator of recent changes.
                    if (-not $wroteAny -and $group.WhenChanged -and ($group.WhenChanged -ge $cutoffDate)) {
                        $results += [PSCustomObject]@{
                            GroupName           = $resolvedGroupName
                            Domain              = $domain
                            AttributeChanged    = 'whenChanged'
                            ChangeTime          = $group.WhenChanged.ToString('yyyy-MM-dd HH:mm:ss')
                            OriginatingDC       = $dc
                            ChangeVersion       = 0
                        }
                    }
                } catch {
                    # Skip individual group errors
                }
            }
        } catch {
            # Skip domain errors
        }
    }

    if ($results.Count -eq 0) {
        $results = @([PSCustomObject]@{
            GroupName        = 'No changes found'
            Domain           = '-'
            AttributeChanged = '-'
            ChangeTime       = '-'
            OriginatingDC    = '-'
            ChangeVersion    = 0
        })
    }

    $results | ConvertTo-Json -Depth 5
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
}
