param(
    [string]$Filter = "all",
    [string]$Format = "json",
    [int]$Limit = 1000
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $ADFilter = switch ($Filter) {
        "enabled"  { { Enabled -eq $true } }
        "disabled" { { Enabled -eq $false } }
        "servers"  { { OperatingSystem -like "*Server*" } }
        "workstations" { { OperatingSystem -notlike "*Server*" } }
        default    { "*" }
    }
    
    $Computers = Get-ADComputer -Filter $ADFilter -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, Enabled, LastLogonDate, WhenCreated, WhenChanged, Description, IPv4Address |
                 Select-Object -First $Limit
    
    $Results = foreach ($Computer in $Computers) {
        [PSCustomObject]@{
            Name              = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            IPAddress         = $Computer.IPv4Address
            OperatingSystem   = $Computer.OperatingSystem
            OSVersion         = $Computer.OperatingSystemVersion
            Enabled           = $Computer.Enabled
            Description       = $Computer.Description
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
