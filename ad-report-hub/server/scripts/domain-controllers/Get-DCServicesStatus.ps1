param(
    [string]$ForestDomain = "",
    [string]$TargetDomain = ""
)

$ErrorActionPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'

# ── Services with well-known TCP ports (checked via fast parallel TCP connect) ──
$PortServices = [ordered]@{
    NTDS     = 389   # LDAP → AD DS
    ADWS     = 9389  # AD Web Services
    DNS      = 53    # DNS Server
    Kdc      = 88    # Kerberos KDC
    Netlogon = 445   # SMB / Netlogon
    RpcSs    = 135   # RPC Endpoint Mapper
}

# ── Services queried via WMI (accurate status, requires elevated access) ──
$AgentServices = @(
    "DFSR",
    "AATPSensor",
    "AATPSensorUpdater",
    "AzureADConnectHealthAgent",
    "CSFalconService",
    "NPSrvHost",
    "Tanium Client",
    "W32Time",
    "Cribl"
)

$CriticalServices = @("NTDS", "Netlogon", "DNS", "Kdc", "ADWS", "DFSR")
$TcpTimeoutMs = 2000

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$serverParam = @{}
if ($ForestDomain)  { $serverParam['Server'] = $ForestDomain }
elseif ($TargetDomain) { $serverParam['Server'] = $TargetDomain }

$MatrixResults = New-Object System.Collections.Generic.List[PSObject]

try {
    $Forest = $null
    try { $Forest = Get-ADForest @serverParam -ErrorAction Stop } catch { $Forest = $null }

    $AllDomains = if ($TargetDomain) { @($TargetDomain) }
                  elseif ($Forest -and $Forest.Domains) { $Forest.Domains }
                  elseif ($ForestDomain) { @($ForestDomain) }
                  else { @() }

    foreach ($Domain in $AllDomains) {
        try {
            $DCs = Get-ADDomainController -Filter * -Server $Domain -ErrorAction SilentlyContinue
            $dcList = @($DCs | Where-Object { $_ -and $_.HostName })
            if ($dcList.Count -eq 0) { continue }

            # ────────────────────────────────────────────────────
            # Phase 1: TCP port probes (all DCs × all port-services, fully parallel)
            # ────────────────────────────────────────────────────
            $totalPortJobs = $dcList.Count * $PortServices.Count
            $pool = [runspacefactory]::CreateRunspacePool(1, [math]::Min(128, $totalPortJobs))
            $pool.Open()

            $portJobs = New-Object System.Collections.Generic.List[PSObject]

            foreach ($DC in $dcList) {
                $fqdn   = [string]$DC.HostName
                $target = if ($DC.IPv4Address) { $DC.IPv4Address } else { $fqdn }

                foreach ($entry in $PortServices.GetEnumerator()) {
                    $ps = [powershell]::Create().AddScript({
                        param($Target, $Port, $TimeoutMs)
                        try {
                            $tcp = New-Object System.Net.Sockets.TcpClient
                            $ar  = $tcp.BeginConnect($Target, $Port, $null, $null)
                            $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
                            if ($ok -and $tcp.Connected) { $tcp.Close(); return "Running" }
                            else { $tcp.Close(); return "Stopped" }
                        } catch { return "Stopped" }
                    }).AddArgument($target).AddArgument($entry.Value).AddArgument($TcpTimeoutMs)

                    $ps.RunspacePool = $pool
                    $portJobs.Add([PSCustomObject]@{
                        DC = $fqdn; Svc = $entry.Key; PS = $ps; H = $ps.BeginInvoke()
                    })
                }
            }

            $portResults = @{}
            foreach ($j in $portJobs) {
                try {
                    $r = $j.PS.EndInvoke($j.H)
                    $v = if ($r -and $r.Count -gt 0) { [string]$r[0] } else { "Stopped" }
                } catch { $v = "Stopped" }
                $j.PS.Dispose()
                if (-not $portResults.ContainsKey($j.DC)) { $portResults[$j.DC] = @{} }
                $portResults[$j.DC][$j.Svc] = $v
            }
            $pool.Close(); $pool.Dispose()

            # ────────────────────────────────────────────────────
            # Phase 2: Agent services via parallel WMI per DC (DCOM, no WinRM needed)
            # ────────────────────────────────────────────────────
            $agentResults = @{}
            # Build WQL filter once
            $wqlParts = $AgentServices | ForEach-Object { "Name='$($_)'" }
            $wqlFilter = $wqlParts -join ' OR '

            $aPool = [runspacefactory]::CreateRunspacePool(1, [math]::Min(32, $dcList.Count))
            $aPool.Open()
            $agentJobs = New-Object System.Collections.Generic.List[PSObject]

            foreach ($DC in $dcList) {
                $fqdn = [string]$DC.HostName
                $ps = [powershell]::Create().AddScript({
                    param($Computer, $Filter, $SvcList)
                    $result = @{}
                    try {
                        $svcs = Get-WmiObject -Class Win32_Service -ComputerName $Computer -Filter $Filter -ErrorAction Stop
                        foreach ($s in @($svcs)) {
                            $result[$s.Name] = [string]$s.State
                        }
                        foreach ($name in $SvcList) {
                            if (-not $result.ContainsKey($name)) {
                                $result[$name] = "Not Installed"
                            }
                        }
                    } catch {
                        $msg = $_.Exception.Message
                        $status = if ($msg -match 'Access is denied') { "Requires Elevated Access" } else { "Error" }
                        foreach ($name in $SvcList) {
                            $result[$name] = $status
                        }
                    }
                    return $result
                }).AddArgument($fqdn).AddArgument($wqlFilter).AddArgument($AgentServices)

                $ps.RunspacePool = $aPool
                $agentJobs.Add([PSCustomObject]@{
                    DC = $fqdn; PS = $ps; H = $ps.BeginInvoke()
                })
            }

            foreach ($j in $agentJobs) {
                try {
                    $r = $j.PS.EndInvoke($j.H)
                    if ($r -and $r.Count -gt 0 -and $r[0] -is [hashtable]) {
                        $agentResults[$j.DC] = $r[0]
                    } else {
                        $agentResults[$j.DC] = @{}
                        foreach ($name in $AgentServices) { $agentResults[$j.DC][$name] = "Error" }
                    }
                } catch {
                    $agentResults[$j.DC] = @{}
                    foreach ($name in $AgentServices) { $agentResults[$j.DC][$name] = "Error" }
                }
                $j.PS.Dispose()
            }
            $aPool.Close(); $aPool.Dispose()

            # ────────────────────────────────────────────────────
            # Build output rows
            # ────────────────────────────────────────────────────
            foreach ($DC in $dcList) {
                $fqdn = [string]$DC.HostName
                $obj = [ordered]@{
                    Domain        = $Domain
                    DCName        = $fqdn
                    Site          = $DC.Site
                    IPAddress     = $DC.IPv4Address
                    IsGC          = $DC.IsGlobalCatalog
                    OSVersion     = $DC.OperatingSystem
                    OverallStatus = "Healthy"
                    QueryMethod   = "PortCheck+WMI"
                }

                $unhealthy = 0
                $pData = $portResults[$fqdn]
                foreach ($entry in $PortServices.GetEnumerator()) {
                    $val = if ($pData -and $pData.ContainsKey($entry.Key)) { $pData[$entry.Key] } else { "Unknown" }
                    $obj[$entry.Key] = $val
                    if ($val -ne "Running" -and $CriticalServices -contains $entry.Key) { $unhealthy++ }
                }

                $aData = $agentResults[$fqdn]
                foreach ($svcName in $AgentServices) {
                    $val = if ($aData -and $aData.ContainsKey($svcName)) { $aData[$svcName] } else { "Unknown" }
                    $obj[$svcName] = $val
                }

                if ($unhealthy -ge 4)     { $obj["OverallStatus"] = "Offline"  }
                elseif ($unhealthy -ge 2) { $obj["OverallStatus"] = "Critical" }
                elseif ($unhealthy -ge 1) { $obj["OverallStatus"] = "Warning"  }

                $MatrixResults.Add([PSCustomObject]$obj)
            }
        }
        catch { }
    }
}
catch {
    @{ Error = "Failed to query forest: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($MatrixResults) | ConvertTo-Json -Depth 5 -Compress
