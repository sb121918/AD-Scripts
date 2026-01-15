$ForestDomains = (Get-ADForest).Domains
$FullReport = ForEach ($Domain in $ForestDomains) {
    Write-Host "Querying domain: $Domain" -ForegroundColor Cyan
    
    # Get all DCs for the current domain in the loop
    $DCs = Get-ADDomainController -Filter * -Server $Domain
    
    ForEach ($DC in $DCs) {
        Write-Host "  Processing: $($DC.HostName)" -ForegroundColor Gray
        
               # Attempting a more compatible CIM connection
        try {
            # DCOM is often better for older servers, WinRM is better for new ones
            $SessionOption = New-CimSessionOption -Protocol Dcom 
            $CimSession = New-CimSession -ComputerName $DC.HostName -Option $SessionOption -ErrorAction Stop -ConnectionTimeoutSec 5
            
            $NetworkConfigs = Get-CimInstance -CimSession $CimSession -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True"
            
            # Combine multiple adapter settings if they exist
            $DnsServers = ($NetworkConfigs.DNSServerSearchOrder | Where-Object {$_} | Select-Object -Unique) -join "; "
            $DnsSearch  = ($NetworkConfigs.DNSDomainSuffixSearchOrder | Where-Object {$_} | Select-Object -Unique) -join "; "
            
            Remove-CimSession $CimSession
        } catch {
            # If CIM fails, variables stay as "Access Denied"
        }

        # Build the final object for this DC
        [PSCustomObject]@{
            "Domain Name"     = $DC.Domain
            "DNS Name"        = $DC.HostName
            "NetbiosName"     = $DC.Name
            "IP Address"      = $DC.IPv4Address
            "AD Site"         = $DC.Site
            "OS Version"      = $DC.OperatingSystem
            "FSMO Roles"      = ($DC.OperationMasterRoles -join ", ")
       
        }
    }
}

# Display the final report in a sortable window
$FullReport | Out-GridView -Title "Multi-Domain Forest DC Audit"