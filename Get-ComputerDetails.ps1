param(
    [Parameter(Mandatory=$true)]
    [string]$ComputerName
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $Computer = Get-ADComputer -Identity $ComputerName -Properties *
    
    if ($null -eq $Computer) {
        @{ Error = "Computer not found" } | ConvertTo-Json
        exit
    }
    
    $Result = [PSCustomObject]@{
        Name                     = $Computer.Name
        DNSHostName              = $Computer.DNSHostName
        SamAccountName           = $Computer.SamAccountName
        SID                      = $Computer.SID.Value
        OperatingSystem          = $Computer.OperatingSystem
        OperatingSystemVersion   = $Computer.OperatingSystemVersion
        OperatingSystemHotfix    = $Computer.OperatingSystemHotfix
        IPv4Address              = $Computer.IPv4Address
        Enabled                  = $Computer.Enabled
        Description              = $Computer.Description
        Location                 = $Computer.Location
        ManagedBy                = if ($Computer.ManagedBy) { ($Computer.ManagedBy -split ',')[0].Replace("CN=","") } else { $null }
        LastLogonDate            = $Computer.LastLogonDate
        PasswordLastSet          = $Computer.PasswordLastSet
        ServicePrincipalNames    = @($Computer.ServicePrincipalName)
        Created                  = $Computer.WhenCreated
        Modified                 = $Computer.WhenChanged
        DistinguishedName        = $Computer.DistinguishedName
    }
    
    $Result | ConvertTo-Json -Depth 4
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
