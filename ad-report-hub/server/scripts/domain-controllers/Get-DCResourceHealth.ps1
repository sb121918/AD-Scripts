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
if ($ForestDomain) { $serverParam['Server'] = $ForestDomain }
elseif ($TargetDomain) { $serverParam['Server'] = $TargetDomain }

$FullReport = New-Object System.Collections.Generic.List[PSObject]

try {
    $Forest = $null
    try {
        $Forest = Get-ADForest @serverParam -ErrorAction Stop
    } catch {
        $Forest = $null
    }
    $AllDomains = if ($TargetDomain) { @($TargetDomain) } elseif ($Forest -and $Forest.Domains) { $Forest.Domains } elseif ($ForestDomain) { @($ForestDomain) } else { @() }

    foreach ($DomainName in $AllDomains) {
        try {
            $DCs = Get-ADDomainController -Filter * -Server $DomainName -ErrorAction SilentlyContinue

            foreach ($DC in $DCs) {
                $Target = $DC.HostName
                
                try {
                    # CPU Usage
                    $CPU = Get-WmiObject -ComputerName $Target -Class Win32_Processor -ErrorAction Stop | 
                           Measure-Object -Property LoadPercentage -Average | 
                           Select-Object -ExpandProperty Average

                    # RAM Usage
                    $OS = Get-WmiObject -ComputerName $Target -Class Win32_OperatingSystem -ErrorAction Stop
                    $TotalRAM = [Math]::Round($OS.TotalVisibleMemorySize / 1MB, 2)
                    $FreeRAM  = [Math]::Round($OS.FreePhysicalMemory / 1MB, 2)
                    $UsedRAM  = $TotalRAM - $FreeRAM
                    $RAMPercent = [Math]::Round(($UsedRAM / $TotalRAM) * 100, 2)

                    # Disk Usage (System Drive C:)
                    $Disk = Get-WmiObject -ComputerName $Target -Class Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
                    $TotalDisk = [Math]::Round($Disk.Size / 1GB, 2)
                    $FreeDisk  = [Math]::Round($Disk.FreeSpace / 1GB, 2)
                    $UsedDisk  = $TotalDisk - $FreeDisk
                    $DiskPercent = [Math]::Round(($UsedDisk / $TotalDisk) * 100, 2)

                    $ReportLine = [PSCustomObject]@{
                        Domain         = $DomainName
                        DCName         = $Target
                        Status         = "Online"
                        'CPU_Load_%'   = $CPU
                        'RAM_Total_GB' = $TotalRAM
                        'RAM_Used_GB'  = $UsedRAM
                        'RAM_Used_%'   = $RAMPercent
                        'C_Total_GB'   = $TotalDisk
                        'C_Free_GB'    = $FreeDisk
                        'C_Used_%'     = $DiskPercent
                        IPAddress      = $DC.IPv4Address
                        Site           = $DC.Site
                    }
                    $FullReport.Add($ReportLine)

                } catch {
                    $ErrorLine = [PSCustomObject]@{
                        Domain         = $DomainName
                        DCName         = $Target
                        Status         = "Unreachable/Offline"
                        'CPU_Load_%'   = "N/A"
                        'RAM_Total_GB' = "N/A"
                        'RAM_Used_GB'  = "N/A"
                        'RAM_Used_%'   = "N/A"
                        'C_Total_GB'   = "N/A"
                        'C_Free_GB'    = "N/A"
                        'C_Used_%'     = "N/A"
                        IPAddress      = $DC.IPv4Address
                        Site           = $DC.Site
                    }
                    $FullReport.Add($ErrorLine)
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

@($FullReport) | ConvertTo-Json -Depth 5
