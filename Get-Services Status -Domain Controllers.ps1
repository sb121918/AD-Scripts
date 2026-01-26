# --- Configuration ---
$ExportToCSV = $true
$FilePath = "$HOME\Desktop\DC_Service_Matrix_Report.csv"

# Your specific list of services
$ServicesToCheck = @(
    "Lsass", "RpcSs", "NTDS", "Kdc", "Netlogon", "EventLog", 
    "DNS", "DFSR", "Cribl", "CSFalconService", "NPSrvHost", 
    "W32Time", "AzureADConnectAgentUpdater"
)

Write-Host "Starting Forest-Wide DC Matrix Check..." -ForegroundColor Cyan

try {
    $AllDomains = (Get-ADForest).Domains
    $MatrixResults = New-Object System.Collections.Generic.List[PSObject]

    foreach ($Domain in $AllDomains) {
        Write-Host "Scanning Domain: $Domain" -ForegroundColor White
        $DCs = Get-ADDomainController -Filter * -Server $Domain

        foreach ($DC in $DCs) {
            Write-Host "  -> Querying $($DC.HostName)" -ForegroundColor Gray
            
            # Create a base object for this DC
            $DCObject = [ordered]@{
                Domain    = $Domain
                DCName    = $DC.HostName
                Site      = $DC.Site
                IPAddress = $DC.IPv4Address
            }

            try {
                # Get the status of all services for this DC in one call
                $ServiceStatuses = Get-Service -Name $ServicesToCheck -ComputerName $DC.HostName -ErrorAction SilentlyContinue
                
                # Add each service status as a column
                foreach ($SvcName in $ServicesToCheck) {
                    $Match = $ServiceStatuses | Where-Object { $_.Name -eq $SvcName }
                    $StatusValue = if ($Match) { $Match.Status } else { "Not Installed" }
                    $DCObject.Add($SvcName, $StatusValue)
                }
            } catch {
                # If connection fails, mark all services as Offline
                foreach ($SvcName in $ServicesToCheck) {
                    $DCObject.Add($SvcName, "OFFLINE/UNREACHABLE")
                }
            }
            
            # Convert the ordered hash to a PSCustomObject and add to list
            $MatrixResults.Add([PSCustomObject]$DCObject)
        }
    }

    # --- Output and Export ---
    if ($MatrixResults.Count -gt 0) {
        # Visual Table Summary (Console)
        $MatrixResults | Format-Table -AutoSize

        # The Grid Output (Searchable/Sortable)
        $MatrixResults | Out-GridView -Title "DC Service Matrix - Forest Health"

        if ($ExportToCSV) {
            $MatrixResults | Export-Csv -Path $FilePath -NoTypeInformation -Encoding utf8
            Write-Host "`nMatrix report generated: $FilePath" -ForegroundColor Green
        }
    }

} catch {
    Write-Host "Critical Error: $($_.Exception.Message)" -ForegroundColor Red
}# --- Status: Ping, Netlogon, NTDS, DNS, Dcdiag Tests ---

#############################################################################
########################### Define Variables ################################

# Set the report path to the current directory
$reportpath = "$PSScriptRoot\AD_Health_Report.html" 
$timeout = "60"

############################### HTML Report Header ############################

$Header = @"
<html>
<head>
    <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
    <title>AD Status Report</title>
    <style type="text/css">
        body { font-family: Tahoma; margin: 10px; }
        table { border-collapse: collapse; width: 100%; border: 1px solid #000; }
        td, th { border: 1px solid #999; padding: 5px; font-size: 11px; text-align: center; }
        th { background-color: IndianRed; color: white; font-weight: bold; }
        .title-row { background-color: Lavender; color: #003399; font-size: 18px; font-weight: bold; }
        .dc-name { background-color: GainsBoro; font-weight: bold; }
        .pass { background-color: Aquamarine; font-weight: bold; }
        .fail { background-color: Red; color: white; font-weight: bold; }
        .timeout { background-color: Yellow; font-weight: bold; }
    </style>
</head>
<body>
    <table>
        <tr class="title-row"><td colspan="10">Active Directory Health Check</td></tr>
        <tr>
            <th>Identity</th>
            <th>Ping Status</th>
            <th>Netlogon Service</th>
            <th>NTDS Service</th>
            <th>DNS Service</th>
            <th>Netlogon Test</th>
            <th>Replication Test</th>
            <th>Services Test</th>
            <th>Advertising Test</th>
            <th>FSMO Check</th>
        </tr>
"@

$Header | Out-File $reportpath -Encoding utf8

##################################### Get ALL DC Servers #################################

Write-Host "Collecting Domain Controllers from the Forest..." -ForegroundColor Cyan
$getForest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()
$DCServers = $getForest.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 

##################################### Health Check Logic #################################

foreach ($DC in $DCServers) {
    $Identity = $DC
    $Row = "<tr><td class='dc-name'>$Identity</td>"
    Write-Host "Checking $DC..." -ForegroundColor White

    if (Test-Connection -ComputerName $DC -Count 1 -Quiet) {
        Write-Host "  Ping: Success" -ForegroundColor Green
        $Row += "<td class='pass'>Success</td>"

        # List of Services and Tests to run via Jobs
        $CheckItems = @(
            @{Name="Netlogon"; Type="Service"},
            @{Name="NTDS"; Type="Service"},
            @{Name="DNS"; Type="Service"},
            @{Name="netlogons"; Type="Dcdiag"},
            @{Name="Replications"; Type="Dcdiag"},
            @{Name="Services"; Type="Dcdiag"},
            @{Name="Advertising"; Type="Dcdiag"},
            @{Name="FsmoCheck"; Type="Dcdiag"}
        )

        foreach ($Item in $CheckItems) {
            $Job = if ($Item.Type -eq "Service") {
                Start-Job -ScriptBlock { Get-Service -ComputerName $args[0] -Name $args[1] -ErrorAction SilentlyContinue } -ArgumentList $DC, $Item.Name
            } else {
                Start-Job -ScriptBlock { 
                    $out = dcdiag /test:$($args[1]) /s:$($args[0])
                    if ($out -match "passed test $($args[1])") { "Passed" } else { "Failed" }
                } -ArgumentList $DC, $Item.Name
            }

            if (Wait-Job $Job -Timeout $timeout) {
                $Result = Receive-Job $Job
                if ($Item.Type -eq "Service") {
                    if ($Result.Status -eq "Running") {
                        $Row += "<td class='pass'>Running</td>"
                    } else {
                        $Row += "<td class='fail'>$($Result.Status)</td>"
                    }
                } else {
                    if ($Result -eq "Passed") {
                        $Row += "<td class='pass'>Passed</td>"
                    } else {
                        $Row += "<td class='fail'>Failed</td>"
                    }
                }
            } else {
                Stop-Job $Job
                $Row += "<td class='timeout'>Timeout</td>"
            }
        }
    } else {
        Write-Host "  Ping: Failed" -ForegroundColor Red
        $Row += "<td class='fail' colspan='9'>Ping Failed / Offline</td>"
    }

    $Row += "</tr>"
    $Row | Out-File $reportpath -Append -Encoding utf8
}

############################################ Finalize ##################################

"</table></body></html>" | Out-File $reportpath -Append -Encoding utf8

Write-Host "`nHealth Check Complete!" -ForegroundColor Cyan
Write-Host "Report saved to: $reportpath" -ForegroundColor Yellow

# Automatically open the report
Invoke-Item $reportpath