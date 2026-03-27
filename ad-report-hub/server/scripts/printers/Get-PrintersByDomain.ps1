param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

$ErrorActionPreference = 'SilentlyContinue'

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

try {
    # Get all printer objects from the specified domain
    $Printers = Get-ADObject -Filter { objectClass -eq 'printQueue' } `
        -Server $TargetDomain `
        -Properties printColor, printDuplexSupported, serverName, location, printerName, uNCName, portName, driverName, printShareName `
        -ErrorAction Stop

    if ($null -eq $Printers -or @($Printers).Count -eq 0) {
        @() | ConvertTo-Json
        exit 0
    }

    $Results = @($Printers | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.printerName
            Server = $_.serverName
            Location = $_.location
            Domain = $TargetDomain
            Color = [bool]$_.printColor
            Duplex = [bool]$_.printDuplexSupported
            UNCPath = $_.uNCName
            PortName = $_.portName
            DriverName = $_.driverName
            ShareName = $_.printShareName
            Published = $true
            DistinguishedName = $_.DistinguishedName
        }
    } | Sort-Object Name)

    @($Results) | ConvertTo-Json -Depth 4
}
catch {
    @{ Error = "Failed to query printers: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}
