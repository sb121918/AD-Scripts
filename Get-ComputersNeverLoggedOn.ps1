param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all computers that never logged on (no limit)
    $Computers = Get-ADComputer -Filter { LastLogonDate -notlike "*" } -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, Enabled, WhenCreated
    
    $Results = foreach ($Computer in $Computers) {
        [PSCustomObject]@{
            Name              = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            OperatingSystem   = $Computer.OperatingSystem
            OSVersion         = $Computer.OperatingSystemVersion
            Enabled           = $Computer.Enabled
            LastLogon         = "Never"
            Created           = $Computer.WhenCreated
            DistinguishedName = $Computer.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
