# --- Configuration ---
$ExportToCSV = $true
$FilePath = "$HOME\Desktop\DC_Resource_Usage_Report.csv"

Write-Host "Gathering Domain Controllers from the entire forest..." -ForegroundColor Cyan

try {
    # 1. Get all domains in the forest
    $AllDomains = (Get-ADForest).Domains
    $FullReport = New-Object System.Collections.Generic.List[PSObject]

    foreach ($DomainName in $AllDomains) {
        Write-Host "Processing Domain: $DomainName" -ForegroundColor White
        
        # 2. Get all DCs in the current domain
        $DCs = Get-ADDomainController -Filter * -Server $DomainName

        foreach ($DC in $DCs) {
            $Target = $DC.HostName
            Write-Host "  Querying $Target..." -ForegroundColor Gray
            
            try {
                # --- CPU Usage ---
                $CPU = Get-WmiObject -ComputerName $Target -Class Win32_Processor | 
                       Measure-Object -Property LoadPercentage -Average | 
                       Select-Object -ExpandProperty Average

                # --- RAM Usage ---
                $OS = Get-WmiObject -ComputerName $Target -Class Win32_OperatingSystem
                $TotalRAM = [Math]::Round($OS.TotalVisibleMemorySize / 1MB, 2)
                $FreeRAM  = [Math]::Round($OS.FreePhysicalMemory / 1MB, 2)
                $UsedRAM  = $TotalRAM - $FreeRAM
                $RAMPercent = [Math]::Round(($UsedRAM / $TotalRAM) * 100, 2)

                # --- Disk Usage (System Drive C:) ---
                $Disk = Get-WmiObject -ComputerName $Target -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
                $TotalDisk = [Math]::Round($Disk.Size / 1GB, 2)
                $FreeDisk  = [Math]::Round($Disk.FreeSpace / 1GB, 2)
                $UsedDisk  = $TotalDisk - $FreeDisk
                $DiskPercent = [Math]::Round(($UsedDisk / $TotalDisk) * 100, 2)

                # --- Build the Report Object ---
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
                # Handle Offline/Unreachable DCs
                $ErrorLine = [PSCustomObject]@{
                    Domain         = $DomainName
                    DCName         = $Target
                    Status         = "Unreachable/Offline"
                    'CPU_Load_%'   = "N/A"
                    'RAM_Total_GB' = "N/A"
                    'RAM_Used_%'   = "N/A"
                    'C_Used_%'     = "N/A"
                }
                $FullReport.Add($ErrorLine)
                Write-Host "    [!] Error: Could not connect to $Target" -ForegroundColor Red
            }
        }
    }

    # --- Output ---
    if ($FullReport.Count -gt 0) {
        # Visual Table Summary
        $FullReport | Out-GridView -Title "Active Directory Forest Resource Monitor"

        if ($ExportToCSV) {
            $FullReport | Export-Csv -Path $FilePath -NoTypeInformation -Encoding utf8
            Write-Host "`nReport saved to: $FilePath" -ForegroundColor Green
        }
    }

} catch {
    Write-Host "Critical Error: $($_.Exception.Message)" -ForegroundColor Red
}