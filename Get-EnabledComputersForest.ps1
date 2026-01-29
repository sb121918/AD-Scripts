param()

$ErrorActionPreference = 'SilentlyContinue'

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$Results = @()

try {
    $Forest = Get-ADForest -ErrorAction Stop
    
    foreach ($DomainName in $Forest.Domains) {
        try {
            $Computers = Get-ADComputer -Filter { Enabled -eq $true } -Server $DomainName -Properties `
                Name, OperatingSystem, OperatingSystemVersion, LastLogonDate, Created, Modified, `
                Description, DNSHostName, Enabled, DistinguishedName -ErrorAction SilentlyContinue
            
            foreach ($Computer in $Computers) {
                $Results += [PSCustomObject]@{
                    Name                   = $Computer.Name
                    DNSHostName            = $Computer.DNSHostName
                    OperatingSystem        = $Computer.OperatingSystem
                    OperatingSystemVersion = $Computer.OperatingSystemVersion
                    Enabled                = $Computer.Enabled
                    LastLogon              = $Computer.LastLogonDate
                    Created                = $Computer.Created
                    Modified               = $Computer.Modified
                    Description            = $Computer.Description
                    Domain                 = $DomainName
                    DistinguishedName      = $Computer.DistinguishedName
                }
            }
        } catch { }
    }
}
catch {
    @{ Error = "Failed to query computers: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($Results) | ConvertTo-Json -Depth 3
