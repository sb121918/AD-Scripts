param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all printers (no limit)
    $Printers = Get-ADObject -Filter { objectClass -eq "printQueue" } -Properties printerName, serverName, portName, printColor, printDuplexSupported, location, Description, WhenCreated
    
    $Results = foreach ($Printer in $Printers) {
        [PSCustomObject]@{
            Name              = $Printer.printerName
            Server            = $Printer.serverName
            Port              = $Printer.portName
            Location          = $Printer.location
            Description       = $Printer.Description
            Color             = $Printer.printColor
            Duplex            = $Printer.printDuplexSupported
            Created           = $Printer.WhenCreated
            DistinguishedName = $Printer.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
