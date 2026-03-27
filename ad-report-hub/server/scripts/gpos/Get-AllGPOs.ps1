param()

try {
    try {
        Import-Module GroupPolicy -ErrorAction Stop
    } catch {
        @{
            Error = "Failed to load GroupPolicy module: $($_.Exception.Message)"
            Remediation = @{
                WindowsClient = "Run: Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0"
                WindowsServer = "Run: Install-WindowsFeature GPMC"
                Notes = "After install, restart the Node server / PowerShell session."
            }
        } | ConvertTo-Json -Depth 6
        exit 0
    }
    
    # Fetch all GPOs (no limit)
    $GPOs = Get-GPO -All
    
    $Results = foreach ($GPO in $GPOs) {
        [PSCustomObject]@{
            Name              = $GPO.DisplayName
            ID                = $GPO.Id
            Status            = $GPO.GpoStatus
            CreationTime      = $GPO.CreationTime
            ModificationTime  = $GPO.ModificationTime
            Owner             = $GPO.Owner
            DomainName        = $GPO.DomainName
            ComputerVersion   = $GPO.Computer.DSVersion
            UserVersion       = $GPO.User.DSVersion
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
