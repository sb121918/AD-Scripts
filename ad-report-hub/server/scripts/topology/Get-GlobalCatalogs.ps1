try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $forest = Get-ADForest
    $results = @()

    foreach ($gcServer in $forest.GlobalCatalogs) {
        try {
            $dc = Get-ADDomainController -Identity $gcServer -ErrorAction SilentlyContinue
            
            $results += [PSCustomObject]@{
                ServerName = $dc.HostName
                Domain = $dc.Domain
                Site = $dc.Site
                IPAddress = ($dc.IPv4Address -join ", ")
                IsGlobalCatalog = $dc.IsGlobalCatalog
                IsReadOnly = $dc.IsReadOnly
                OperatingSystem = $dc.OperatingSystem
                OperatingSystemVersion = $dc.OperatingSystemVersion
            }
        } catch {
            $results += [PSCustomObject]@{
                ServerName = $gcServer
                Domain = "Unknown"
                Site = "Unknown"
                IPAddress = "Unknown"
                IsGlobalCatalog = $true
                IsReadOnly = $false
                OperatingSystem = "Unknown"
                OperatingSystemVersion = "Unknown"
            }
        }
    }

    $results | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
