param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

$ErrorActionPreference = 'SilentlyContinue'

try {
    Import-Module GroupPolicy -ErrorAction Stop
}
catch {
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

$Results = @()

try {
    # Get all GPOs from the specified domain
    $GPOs = Get-GPO -All -Domain $TargetDomain -ErrorAction Stop
    
    if ($null -eq $GPOs -or @($GPOs).Count -eq 0) {
        @($Results) | ConvertTo-Json -Depth 3
        exit 0
    }
    
    foreach ($GPO in $GPOs) {
        # Determine GPO status
        $Status = "Enabled"
        if ($GPO.GpoStatus -eq "AllSettingsDisabled") {
            $Status = "All Settings Disabled"
        } elseif ($GPO.GpoStatus -eq "UserSettingsDisabled") {
            $Status = "User Settings Disabled"
        } elseif ($GPO.GpoStatus -eq "ComputerSettingsDisabled") {
            $Status = "Computer Settings Disabled"
        }
        
        $Results += [PSCustomObject]@{
            Name             = $GPO.DisplayName
            Id               = $GPO.Id.ToString()
            Domain           = $TargetDomain
            Status           = $Status
            GpoStatus        = $GPO.GpoStatus.ToString()
            Owner            = $GPO.Owner
            CreationTime     = $GPO.CreationTime
            ModificationTime = $GPO.ModificationTime
            Description      = $GPO.Description
        }
    }
}
catch {
    @{ Error = "Failed to query GPOs for ${TargetDomain}: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($Results) | ConvertTo-Json -Depth 3
