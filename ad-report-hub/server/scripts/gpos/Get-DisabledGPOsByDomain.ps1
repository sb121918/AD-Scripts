param(
    [Parameter(Mandatory=$false)]
    [string]$TargetDomain = ""
)

$ErrorActionPreference = 'SilentlyContinue'

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

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        @{
            Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)"
            Remediation = @{
                WindowsClient = "Run: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
                WindowsServer = "Run: Install-WindowsFeature RSAT-AD-PowerShell"
                Notes = "After install, restart the Node server / PowerShell session."
            }
        } | ConvertTo-Json -Depth 6
        exit 0
    }
}
catch {
    @{ Error = "Failed to load required modules: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$Results = @()

try {
    # Determine which domains to query
    $DomainsToQuery = @()
    
    if ($TargetDomain -and $TargetDomain.Trim() -ne "") {
        $DomainsToQuery = @($TargetDomain.Trim())
    } else {
        # Query all domains in forest
        $Forest = Get-ADForest -ErrorAction Stop
        $DomainsToQuery = $Forest.Domains
    }
    
    foreach ($DomainName in $DomainsToQuery) {
        try {
            # Get all GPOs from this domain
            $GPOs = Get-GPO -All -Domain $DomainName -ErrorAction SilentlyContinue
            
            foreach ($GPO in $GPOs) {
                # Check if any settings are disabled
                if ($GPO.GpoStatus -ne "AllSettingsEnabled") {
                    # Determine GPO status
                    $Status = "Unknown"
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
                        Domain           = $DomainName
                        Status           = $Status
                        GpoStatus        = $GPO.GpoStatus.ToString()
                        Owner            = $GPO.Owner
                        CreationTime     = $GPO.CreationTime
                        ModificationTime = $GPO.ModificationTime
                        Description      = $GPO.Description
                    }
                }
            }
        } catch { }
    }
}
catch {
    @{ Error = "Failed to query disabled GPOs: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($Results) | ConvertTo-Json -Depth 3
