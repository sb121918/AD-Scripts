param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $Forest = Get-ADForest
    $ForestRoot = $Forest.RootDomain

    # Define the Search Criteria (EA6 or SVC tag)
    $LdapFilter = "(|(extensionAttribute6=*Non-Human Service Account*)(employeeID=*SVC*)(employeeNumber=*SVC*))"

    # Query the Global Catalog (Port 3268)
    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server "${ForestRoot}:3268" -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, whenChanged, Enabled, Manager

    # Filter for accounts where Manager is null or empty (no limit)
    $OrphanedAccounts = $Users | Where-Object { [string]::IsNullOrWhiteSpace($_.Manager) }

    $Results = foreach ($User in $OrphanedAccounts) {
        # Clean Domain Parsing
        $DnParts = $User.DistinguishedName.Split(',') | Where-Object { $_ -like "DC=*" }
        $CleanDomain = ($DnParts -replace "DC=", "") -join "."

        [PSCustomObject]@{
            Name              = $User.Name
            SamAccountName    = $User.SamAccountName
            AccountType       = "Service Account (Orphaned)"
            Enabled           = $User.Enabled
            Domain            = $CleanDomain
            Created           = $User.whenCreated
            LastModified      = $User.whenChanged
            EA6_Value         = $User.extensionAttribute6
            EmployeeID        = $User.employeeID
            EmployeeNumber    = $User.employeeNumber
            ManagerStatus     = "NOT ASSIGNED"
            Title             = $User.Title
            Department        = $User.Department
            DistinguishedName = $User.DistinguishedName
        }
    }

    @($Results) | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
