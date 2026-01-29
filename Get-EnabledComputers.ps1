param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all enabled computers (no limit)
    $Computers = Get-ADComputer -Filter { Enabled -eq $true } -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, LastLogonDate, WhenCreated
    
    $Results = foreach ($Computer in $Computers) {
        [PSCustomObject]@{
            Name              = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            OperatingSystem   = $Computer.OperatingSystem
            OSVersion         = $Computer.OperatingSystemVersion
            Enabled           = $true
            LastLogon         = $Computer.LastLogonDate
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
