param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $Forest = Get-ADForest
    $ForestRoot = $Forest.RootDomain

    # LDAP Filter: EA6 contains "Service Account" OR EmployeeID contains "SVC" OR EmployeeNumber contains "SVC"
    $LdapFilter = "(|(extensionAttribute6=*Service Account*)(employeeID=*SVC*)(employeeNumber=*SVC*))"

    # Query the Global Catalog (Port 3268) for forest-wide results
    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server "${ForestRoot}:3268" -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled, Manager

    $Results = foreach ($User in $Users) {
        # Clean Domain Parsing from DN
        $DnParts = $User.DistinguishedName.Split(',') | Where-Object { $_ -like "DC=*" }
        $CleanDomain = ($DnParts -replace "DC=", "") -join "."
        
        $ManagerName = $null
        $ManagerStatus = "NOT ASSIGNED"
        if ($User.Manager) {
            $ManagerStatus = "Assigned"
            try {
                # Extract CN from Manager DN
                if ($User.Manager -match "^CN=([^,]+),") {
                    $ManagerName = $Matches[1]
                }
            } catch {
                $ManagerName = "Unable to resolve"
            }
        }

        [PSCustomObject]@{
            Name              = $User.Name
            SamAccountName    = $User.SamAccountName
            AccountType       = "Service Account"
            Enabled           = $User.Enabled
            Domain            = $CleanDomain
            Created           = $User.whenCreated
            EA6_Value         = $User.extensionAttribute6
            EmployeeID        = $User.employeeID
            EmployeeNumber    = $User.employeeNumber
            ManagerStatus     = $ManagerStatus
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
