param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    # LDAP Filter: EA6 contains "Service Account" OR EmployeeID contains "SVC" OR EmployeeNumber contains "SVC"
    $LdapFilter = "(|(extensionAttribute6=*Service Account*)(employeeID=*SVC*)(employeeNumber=*SVC*))"

    # Fetch all service accounts (no limit)
    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server $TargetDomain -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled, Manager

    $Results = foreach ($User in $Users) {
        $ManagerName = $null
        if ($User.Manager) {
            try {
                $ManagerName = (Get-ADUser -Identity $User.Manager -Server $TargetDomain).Name
            } catch {
                $ManagerName = "Unable to resolve"
            }
        }

        [PSCustomObject]@{
            Name              = $User.Name
            SamAccountName    = $User.SamAccountName
            AccountType       = "Service Account"
            Enabled           = $User.Enabled
            Domain            = $TargetDomain
            Created           = $User.whenCreated
            EA6_Value         = $User.extensionAttribute6
            EmployeeID        = $User.employeeID
            EmployeeNumber    = $User.employeeNumber
            ManagerStatus     = if ($User.Manager) { "Assigned" } else { "NOT ASSIGNED" }
            ManagerName       = $ManagerName
            Title             = $User.Title
            Department        = $User.Department
            Description       = $User.Description
            DistinguishedName = $User.DistinguishedName
        }
    }

    @($Results) | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
