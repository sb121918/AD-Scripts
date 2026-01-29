param(
    [int]$Days = 60
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $CutoffDate = (Get-Date).AddDays(-$Days)
    
    # Fetch all computers not logged on in X days (no limit)
    $Computers = Get-ADComputer -Filter { LastLogonDate -lt $CutoffDate -and Enabled -eq $true } -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, Enabled, LastLogonDate, WhenCreated |
                 Sort-Object LastLogonDate
    
    $Results = foreach ($Computer in $Computers) {
        $DaysSinceLogon = if ($Computer.LastLogonDate) { ((Get-Date) - $Computer.LastLogonDate).Days } else { "N/A" }
        
        [PSCustomObject]@{
            Name              = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            OperatingSystem   = $Computer.OperatingSystem
            Enabled           = $Computer.Enabled
            LastLogon         = $Computer.LastLogonDate
            DaysSinceLogon    = $DaysSinceLogon
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
