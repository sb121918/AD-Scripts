param(
    [Parameter(Mandatory=$true)]
    [string]$GPOName
)

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

    $GPO = Get-GPO -Name $GPOName -ErrorAction Stop
    
    if ($null -eq $GPO) {
        @{ Error = "GPO not found" } | ConvertTo-Json
        exit
    }
    
    # Get all OUs and domain
    $Domain = Get-ADDomain
    $OUs = Get-ADOrganizationalUnit -Filter * -Properties gpLink
    
    $Links = @()
    
    # Check domain level
    $DomainGpLink = (Get-ADObject -Identity $Domain.DistinguishedName -Properties gpLink).gpLink
    if ($DomainGpLink -match $GPO.Id) {
        $Links += [PSCustomObject]@{
            Target = $Domain.DistinguishedName
            Type   = "Domain"
            Enabled = $true
        }
    }
    
    # Check OUs
    foreach ($OU in $OUs) {
        if ($OU.gpLink -match $GPO.Id) {
            $Links += [PSCustomObject]@{
                Target  = $OU.DistinguishedName
                Type    = "OU"
                Name    = $OU.Name
                Enabled = $true
            }
        }
    }
    
    @{
        GPOName = $GPO.DisplayName
        GPOID   = $GPO.Id
        LinkCount = $Links.Count
        Links   = @($Links)
    } | ConvertTo-Json -Depth 4
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
