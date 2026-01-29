param(
    [Parameter(Mandatory=$true)]
    [string]$OperatingSystem
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all computers by OS (no limit)
    $Computers = Get-ADComputer -Filter "OperatingSystem -like '*$OperatingSystem*'" -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, Enabled, LastLogonDate, WhenCreated
    
    $Results = foreach ($Computer in $Computers) {
        [PSCustomObject]@{
            Name              = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            OperatingSystem   = $Computer.OperatingSystem
            OSVersion         = $Computer.OperatingSystemVersion
            Enabled           = $Computer.Enabled
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
