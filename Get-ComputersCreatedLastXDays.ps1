param(
    [int]$Days = 30
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $CutoffDate = (Get-Date).AddDays(-$Days)
    
    # Fetch all computers created in last X days (no limit)
    $Computers = Get-ADComputer -Filter { WhenCreated -ge $CutoffDate } -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, Enabled, LastLogonDate, WhenCreated |
                 Sort-Object WhenCreated -Descending
    
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
