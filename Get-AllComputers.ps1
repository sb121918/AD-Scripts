param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all computers (no limit)
    $Computers = Get-ADComputer -Filter * -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, Enabled, LastLogonDate, WhenCreated, Description
    
    $Results = foreach ($Computer in $Computers) {
        [PSCustomObject]@{
            Name              = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            OperatingSystem   = $Computer.OperatingSystem
            OSVersion         = $Computer.OperatingSystemVersion
            Enabled           = $Computer.Enabled
            LastLogon         = $Computer.LastLogonDate
            Created           = $Computer.WhenCreated
            Description       = $Computer.Description
            DistinguishedName = $Computer.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
