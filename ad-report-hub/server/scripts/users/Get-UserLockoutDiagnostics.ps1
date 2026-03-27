param(
    [Parameter(Mandatory=$true)]
    [Alias('SearchValue')]
    [string]$Identity,

    [Parameter(Mandatory=$false)]
    [int]$LookbackHours = 24,

    [Parameter(Mandatory=$false)]
    [ValidateSet('PDC','All')]
    [string]$DcScope = 'PDC',

    [Parameter(Mandatory=$false)]
    [int]$MaxEvents = 200,

    [Parameter(Mandatory=$false)]
    [int]$CorrelationWindowMinutes = 5,

    [Parameter(Mandatory=$false)]
    [string]$TargetDomain = ""
)

function To-IsoString([object]$Value) {
    if ($null -eq $Value) { return $null }
    try {
        if ($Value -is [DateTime]) { return $Value.ToString('o') }
        $dt = [DateTime]$Value
        return $dt.ToString('o')
    } catch {
        return [string]$Value
    }
}

function Escape-ADFilterValue([string]$Value) {
    if ($null -eq $Value) { return "" }
    return $Value.Replace("'", "''")
}

function Get-EventDataValue([xml]$Xml, [string]$Name) {
    try {
        $nodes = @($Xml.Event.EventData.Data)
        foreach ($n in $nodes) {
            if ($n.Name -eq $Name) { return [string]$n.'#text' }
        }
        return $null
    } catch {
        return $null
    }
}

function New-Warning([string]$Message, [string]$Dc = $null) {
    if ($Dc) { return "[$Dc] $Message" }
    return $Message
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $warnings = New-Object System.Collections.Generic.List[string]

    $ServerArgs = @{}
    if ($TargetDomain -and $TargetDomain.Trim() -ne "") {
        $ServerArgs.Server = $TargetDomain.Trim()
    }

    # Resolve user
    $id = $Identity.Trim()
    if (-not $id) {
        @{ Error = "Identity is required" } | ConvertTo-Json
        exit 1
    }

    $user = $null
    try {
        $user = Get-ADUser @ServerArgs -Identity $id -Properties LockedOut, LockoutTime, badPwdCount, LastBadPasswordAttempt, UserPrincipalName, mail -ErrorAction Stop
    } catch {
        $escaped = Escape-ADFilterValue $id
        $Filter = "SamAccountName -eq '$escaped' -or UserPrincipalName -eq '$escaped' -or mail -eq '$escaped'"
        $user = Get-ADUser @ServerArgs -Filter $Filter -Properties LockedOut, LockoutTime, badPwdCount, LastBadPasswordAttempt, UserPrincipalName, mail -ErrorAction SilentlyContinue |
                Select-Object -First 1
    }

    if (-not $user) {
        @{ Error = "User not found for identity '$id'" } | ConvertTo-Json
        exit 1
    }

    $domainToUse = $null
    if ($ServerArgs.Server) {
        $domainToUse = $ServerArgs.Server
    } else {
        try {
            $domainToUse = ($user.DistinguishedName -split ',DC=' | Select-Object -Skip 1) -join '.'
        } catch {
            $domainToUse = (Get-ADDomain).DNSRoot
        }
    }

    # AD properties
    $lockoutTime = $null
    try {
        if ($user.LockoutTime -and [int64]$user.LockoutTime -gt 0) {
            $lockoutTime = [DateTime]::FromFileTime([int64]$user.LockoutTime)
        }
    } catch { }

    $adProps = [PSCustomObject]@{
        isLocked              = [bool]$user.LockedOut
        lockoutTime           = To-IsoString $lockoutTime
        lastBadPasswordAttempt = To-IsoString $user.LastBadPasswordAttempt
        badPwdCount           = $user.badPwdCount
    }

    # Determine PDC
    $pdc = $null
    try {
        $pdc = (Get-ADDomain -Server $domainToUse -ErrorAction Stop).PDCEmulator
    } catch {
        try {
            $pdc = (Get-ADDomain -ErrorAction Stop).PDCEmulator
        } catch {
            $warnings.Add((New-Warning "Failed to determine PDC emulator: $($_.Exception.Message)"))
        }
    }

    # Determine DC list for correlation
    $dcs = @()
    if ($DcScope -eq 'PDC') {
        if ($pdc) { $dcs = @($pdc) }
    } else {
        try {
            $dcs = @(Get-ADDomainController -Filter * -Server $domainToUse -ErrorAction Stop | Select-Object -ExpandProperty HostName)
        } catch {
            $warnings.Add((New-Warning "Failed to list domain controllers: $($_.Exception.Message)"))
            if ($pdc) { $dcs = @($pdc) }
        }
    }

    $startTime = (Get-Date).AddHours(-1 * [Math]::Max(1, $LookbackHours))

    # Query 4740 (Account locked out) on PDC (best source for caller computer)
    $lockoutEvents = @()
    if ($pdc) {
        try {
            $lockoutEvents = @(Get-WinEvent -ComputerName $pdc -FilterHashtable @{ LogName='Security'; Id=4740; StartTime=$startTime } -MaxEvents $MaxEvents -ErrorAction Stop)
        } catch {
            $warnings.Add((New-Warning "Failed to query Security log (4740). Ensure access to Remote Event Log/WinRM and firewall rules." $pdc))
        }
    } else {
        $warnings.Add((New-Warning "PDC emulator unknown; skipping 4740 query."))
    }

    $sam = [string]$user.SamAccountName
    $lockouts = New-Object System.Collections.Generic.List[object]

    foreach ($evt in $lockoutEvents) {
        try {
            $xml = [xml]$evt.ToXml()
            $targetUser = Get-EventDataValue -Xml $xml -Name 'TargetUserName'
            if (-not $targetUser) { continue }
            if ($targetUser -ne $sam -and $targetUser -ne $id) { continue }

            $caller = Get-EventDataValue -Xml $xml -Name 'CallerComputerName'
            $evtTime = $evt.TimeCreated

            $windowStart = $evtTime.AddMinutes(-1 * [Math]::Max(1, $CorrelationWindowMinutes))
            $windowEnd = $evtTime.AddMinutes([Math]::Max(1, $CorrelationWindowMinutes))

            $candidates = New-Object System.Collections.Generic.List[object]

            foreach ($dc in $dcs) {
                try {
                    $fails = @(Get-WinEvent -ComputerName $dc -FilterHashtable @{ LogName='Security'; Id=4625,4771,4776; StartTime=$windowStart; EndTime=$windowEnd } -MaxEvents $MaxEvents -ErrorAction Stop)
                    foreach ($f in $fails) {
                        $fx = [xml]$f.ToXml()
                        $uName = Get-EventDataValue -Xml $fx -Name 'TargetUserName'
                        if (-not $uName) { continue }
                        if ($uName -ne $sam -and $uName -ne $id) { continue }

                        $ip = $null
                        $ws = $null
                        if ($f.Id -eq 4625) {
                            $ip = Get-EventDataValue -Xml $fx -Name 'IpAddress'
                            $ws = Get-EventDataValue -Xml $fx -Name 'WorkstationName'
                        } elseif ($f.Id -eq 4771) {
                            $ip = Get-EventDataValue -Xml $fx -Name 'IpAddress'
                            $ws = Get-EventDataValue -Xml $fx -Name 'WorkstationName'
                        } elseif ($f.Id -eq 4776) {
                            $ws = Get-EventDataValue -Xml $fx -Name 'Workstation'
                        }

                        $candidates.Add([PSCustomObject]@{
                            time        = To-IsoString $f.TimeCreated
                            dc          = $dc
                            eventId     = $f.Id
                            ipAddress   = $ip
                            workstation = $ws
                        })
                    }
                } catch {
                    $warnings.Add((New-Warning "Failed to query Security log for correlation events (4625/4771/4776): $($_.Exception.Message)" $dc))
                }
            }

            # Pick best candidate: prefer ones with a real IP, then closest in time
            $best = $null
            if ($candidates.Count -gt 0) {
                $realIp = $candidates | Where-Object { $_.ipAddress -and $_.ipAddress -ne '-' -and $_.ipAddress -ne '::1' -and $_.ipAddress -ne '127.0.0.1' }
                $pool = if ($realIp) { $realIp } else { $candidates }
                $best = $pool |
                    Sort-Object @{
                        Expression = {
                            try {
                                [Math]::Abs((([DateTime]$_.time) - $evtTime).TotalSeconds)
                            } catch { 999999 }
                        }
                    } |
                    Select-Object -First 1
            }

            $lockouts.Add([PSCustomObject]@{
                time           = To-IsoString $evtTime
                dc             = $pdc
                callerComputer = $caller
                ipAddress      = if ($best) { $best.ipAddress } else { $null }
                workstation    = if ($best) { $best.workstation } else { $null }
                eventIdMatched = if ($best) { $best.eventId } else { $null }
                correlatedDc   = if ($best) { $best.dc } else { $null }
                correlatedTime = if ($best) { $best.time } else { $null }
                candidatesCount = $candidates.Count
            })
        } catch {
            $warnings.Add((New-Warning "Failed to parse/correlate a 4740 event: $($_.Exception.Message)" $pdc))
        }
    }

    $result = [PSCustomObject]@{
        identity            = $id
        samAccountName      = $user.SamAccountName
        userPrincipalName   = $user.UserPrincipalName
        distinguishedName   = $user.DistinguishedName
        domain              = $domainToUse
        pdcEmulator         = $pdc
        ad                 = $adProps
        lockouts            = @($lockouts)
        warnings            = @($warnings)
    }

    $result | ConvertTo-Json -Depth 6
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
