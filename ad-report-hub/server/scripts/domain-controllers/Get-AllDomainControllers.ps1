param(
    [string]$ForestDomain = "",
    [string]$TargetDomain = ""
)

$ErrorActionPreference = 'SilentlyContinue'

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$serverParam = @{}
# ForestDomain anchors Get-ADForest discovery to the selected forest
if ($ForestDomain) { $serverParam['Server'] = $ForestDomain }
elseif ($TargetDomain) { $serverParam['Server'] = $TargetDomain }

$AllDCs = @()

try {
    $Forest = $null
    try {
        $Forest = Get-ADForest @serverParam -ErrorAction Stop
    } catch {
        # If forest discovery fails (common when querying a trusted forest without root access),
        # fall back to querying the provided domain only.
        $Forest = $null
    }

    if (-not $Forest) {
        # Minimal object so FSMO checks don't error
        $Forest = [PSCustomObject]@{ SchemaMaster = $null; DomainNamingMaster = $null; Domains = @() }
    }

    # If TargetDomain is provided, run per-domain. Otherwise, run forest-wide.
    # If forest discovery failed, fall back to ForestDomain (single-domain).
    $domainsToQuery = if ($TargetDomain) { @($TargetDomain) } elseif ($Forest.Domains -and $Forest.Domains.Count -gt 0) { $Forest.Domains } elseif ($ForestDomain) { @($ForestDomain) } else { @() }
    
    foreach ($DomainName in $domainsToQuery) {
        try {
            $Domain = Get-ADDomain -Server $DomainName -ErrorAction SilentlyContinue
            $DCs = Get-ADDomainController -Filter * -Server $DomainName -ErrorAction SilentlyContinue
            
            if ($DCs) {
                foreach ($DC in $DCs) {
                    $ComputerObj = $null
                    try {
                        $ComputerObj = Get-ADComputer -Identity $DC.Name -Server $DomainName -Properties OperatingSystem, OperatingSystemVersion, Description, WhenCreated, LastLogonDate -ErrorAction SilentlyContinue
                    } catch { }
                    
                    # Determine FSMO roles held by this DC
                    $FSMORoles = @()
                    if ($Domain -and $Domain.PDCEmulator -eq $DC.HostName) { $FSMORoles += "PDC" }
                    if ($Domain -and $Domain.RIDMaster -eq $DC.HostName) { $FSMORoles += "RID" }
                    if ($Domain -and $Domain.InfrastructureMaster -eq $DC.HostName) { $FSMORoles += "Infrastructure" }
                    if ($Forest.SchemaMaster -eq $DC.HostName) { $FSMORoles += "Schema" }
                    if ($Forest.DomainNamingMaster -eq $DC.HostName) { $FSMORoles += "Naming" }
                    
                    $uptime = $null
                    try {
                        $os = Get-WmiObject Win32_OperatingSystem -ComputerName $DC.HostName -ErrorAction SilentlyContinue
                        if ($os -and $os.LastBootUpTime) {
                            $bootTime = $os.ConvertToDateTime($os.LastBootUpTime)
                            $span = (Get-Date) - $bootTime
                            $uptime = "{0}d {1}h {2}m" -f $span.Days, $span.Hours, $span.Minutes
                        }
                    } catch { }

                    $AllDCs += [PSCustomObject]@{
                        Domain              = $DomainName
                        DCName              = $DC.Name
                        HostName            = $DC.HostName
                        IPv4Address         = $DC.IPv4Address
                        Site                = $DC.Site
                        OperatingSystem     = if ($ComputerObj) { $ComputerObj.OperatingSystem } else { $null }
                        OSVersion           = if ($ComputerObj) { $ComputerObj.OperatingSystemVersion } else { $null }
                        IsGlobalCatalog     = $DC.IsGlobalCatalog
                        IsReadOnly          = $DC.IsReadOnly
                        Enabled             = $DC.Enabled
                        Uptime              = $uptime
                        FSMORoles           = if ($FSMORoles.Count -gt 0) { $FSMORoles -join ", " } else { "None" }
                        Description         = if ($ComputerObj) { $ComputerObj.Description } else { $null }
                        Created             = if ($ComputerObj) { $ComputerObj.WhenCreated } else { $null }
                        LastLogon           = if ($ComputerObj) { $ComputerObj.LastLogonDate } else { $null }
                        DistinguishedName   = $DC.ComputerObjectDN
                    }
                }
            }
        }
        catch { }
    }
}
catch {
    @{ Error = "Failed to query forest: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($AllDCs) | ConvertTo-Json -Depth 5
