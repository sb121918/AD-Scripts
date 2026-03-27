param(
    [int]$Limit = 0,
    [string]$TargetDomain = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $WarningPreference = 'SilentlyContinue'
    $ErrorActionPreference = 'SilentlyContinue'

    function Get-CnFromDn {
        param([string]$Dn)
        if (-not $Dn) { return $null }
        if ($Dn -match '^CN=([^,]+),') { return $matches[1] }
        return $Dn
    }

    $forest = $null
    try {
        if ($TargetDomain) { $forest = Get-ADForest -Server $TargetDomain -ErrorAction Stop }
        else { $forest = Get-ADForest -ErrorAction Stop }
    } catch { $forest = $null }

    if (-not $forest -or -not $forest.Domains) {
        @{ Error = "Failed to discover forest domains. Provide a valid domain in TargetDomain." } | ConvertTo-Json
        exit 1
    }

    $AllDomains = @($forest.Domains)
    $MasterUserList = New-Object System.Collections.Generic.List[PSObject]

    $ADProps = @(
        "extensionAttribute6", 
        "Manager", 
        "EmailAddress", 
        "PasswordLastSet", 
        "WhenCreated",
        "Department",
        "Title"
    )

    foreach ($Domain in $AllDomains) {
        try {
            $DomainUsers = Get-ADUser -LDAPFilter '(!(extensionAttribute6=*))' -Server $Domain `
                -ResultPageSize 2000 -ResultSetSize $null `
                -Properties $ADProps

            foreach ($User in $DomainUsers) {
                $ManagerName = Get-CnFromDn $User.Manager

                # Get OU from DN
                $OU = ($User.DistinguishedName -split ',', 2)[1]

                $MasterUserList.Add([PSCustomObject]@{
                    Name              = $User.Name
                    SamAccountName    = $User.SamAccountName
                    Email             = $User.EmailAddress
                    Manager           = $ManagerName
                    Department        = $User.Department
                    Title             = $User.Title
                    OU                = $OU
                    Domain            = $Domain
                    PasswordLastSet   = $User.PasswordLastSet
                    WhenCreated       = if ($User.WhenCreated) { $User.WhenCreated.ToString('yyyy-MM-dd') } else { $null }
                    EA6               = $User.extensionAttribute6
                    DistinguishedName = $User.DistinguishedName
                })

                if ($Limit -gt 0 -and $MasterUserList.Count -ge $Limit) { break }
            }
        } catch {
            # Skip domain if error
        }
        
        if ($Limit -gt 0 -and $MasterUserList.Count -ge $Limit) { break }
    }

    @($MasterUserList) | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
