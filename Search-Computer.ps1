param(
    [Parameter(Mandatory=$true)]
    [string]$SearchValue
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $WildcardSearch = "*$($SearchValue.Trim('*'))*"
    
    # Search computers (no limit)
    $Computers = Get-ADComputer -Filter "Name -like '$WildcardSearch'" -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, Enabled, LastLogonDate, WhenCreated, Description, IPv4Address
    
    $Results = foreach ($Computer in $Computers) {
        [PSCustomObject]@{
            Name              = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            IPAddress         = $Computer.IPv4Address
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
