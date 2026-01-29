param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain,
    
    [Parameter(Mandatory=$true)]
    [string]$OperatingSystem
)

$ErrorActionPreference = 'SilentlyContinue'

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$Results = @()

try {
    # Build filter for operating system (using wildcard match)
    $OSFilter = "*$OperatingSystem*"
    
    $Computers = Get-ADComputer -Filter { OperatingSystem -like $OSFilter } -Server $TargetDomain -Properties `
        Name, OperatingSystem, OperatingSystemVersion, LastLogonDate, Created, Modified, `
        Description, DNSHostName, Enabled, DistinguishedName -ErrorAction Stop
    
    foreach ($Computer in $Computers) {
        $Results += [PSCustomObject]@{
            Name                   = $Computer.Name
            DNSHostName            = $Computer.DNSHostName
            OperatingSystem        = $Computer.OperatingSystem
            OperatingSystemVersion = $Computer.OperatingSystemVersion
            Enabled                = $Computer.Enabled
            LastLogon              = $Computer.LastLogonDate
            Created                = $Computer.Created
            Modified               = $Computer.Modified
            Description            = $Computer.Description
            Domain                 = $TargetDomain
            DistinguishedName      = $Computer.DistinguishedName
        }
    }
}
catch {
    @{ Error = "Failed to query computers for ${TargetDomain}: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($Results) | ConvertTo-Json -Depth 3
