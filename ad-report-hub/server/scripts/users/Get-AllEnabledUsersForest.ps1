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
        if ($TargetDomain) {
            $forest = Get-ADForest -Server $TargetDomain -ErrorAction Stop
        } else {
            $forest = Get-ADForest -ErrorAction Stop
        }
    } catch {
        $forest = $null
    }

    if (-not $forest -or -not $forest.Domains) {
        @{ Error = "Failed to discover forest domains. Provide a valid domain in TargetDomain." } | ConvertTo-Json
        exit 1
    }

    $ForestDomains = @($forest.Domains)
    $MasterResults = New-Object System.Collections.Generic.List[PSObject]

    foreach ($Domain in $ForestDomains) {
        try {
            $Users = Get-ADUser -Filter 'Enabled -eq $true' -Server $Domain `
              -ResultPageSize 2000 -ResultSetSize $null `
              -Properties `
                GivenName, Surname, DisplayName, Title, UserPrincipalName, SamAccountName, `
                EmailAddress, Department, EmployeeID, EmployeeNumber, EmployeeType, `
                Description, Manager, whenCreated, PasswordNeverExpires, PasswordExpired, `
                extensionAttribute5, extensionAttribute6

            foreach ($User in $Users) {
                # Avoid per-user lookups (too slow for whole-forest). Extract CN from Manager DN.
                $ManagerName = Get-CnFromDn $User.Manager

                $MasterResults.Add([PSCustomObject]@{
                    FirstName            = $User.GivenName
                    LastName             = $User.Surname
                    DisplayName          = $User.DisplayName
                    Title                = $User.Title
                    UPN                  = $User.UserPrincipalName
                    SamAccountName       = $User.SamAccountName
                    Email                = $User.EmailAddress
                    Department           = $User.Department
                    EmployeeID           = $User.EmployeeID
                    EmployeeNumber       = $User.EmployeeNumber
                    EmployeeType         = $User.EmployeeType
                    Description          = $User.Description
                    ManagerName          = $ManagerName
                    Created              = $User.whenCreated
                    PasswordNeverExpires = $User.PasswordNeverExpires
                    PasswordExpired      = $User.PasswordExpired
                    EA5                  = $User.extensionAttribute5
                    EA6                  = $User.extensionAttribute6
                    Domain               = $Domain
                })

                if ($Limit -gt 0 -and $MasterResults.Count -ge $Limit) { break }
            }
        } catch {
            # Skip domain if error
        }
        
        if ($Limit -gt 0 -and $MasterResults.Count -ge $Limit) { break }
    }

    @($MasterResults) | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
