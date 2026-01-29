param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all disabled computers (no limit)
    $Computers = Get-ADComputer -Filter { Enabled -eq $false } -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, LastLogonDate, WhenCreated, WhenChanged
    
    $Results = foreach ($Computer in $Computers) {
        [PSCustomObject]@{
            Name              = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            OperatingSystem   = $Computer.OperatingSystem
            OSVersion         = $Computer.OperatingSystemVersion
            Enabled           = $false
            LastLogon         = $Computer.LastLogonDate
            Created           = $Computer.WhenCreated
            Modified          = $Computer.WhenChanged
            DistinguishedName = $Computer.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
