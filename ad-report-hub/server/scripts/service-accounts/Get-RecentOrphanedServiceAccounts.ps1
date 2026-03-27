param(
    [int]$Days = 365
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $Forest = Get-ADForest
    $ForestRoot = $Forest.RootDomain

    # Calculate LDAP Date
    $ThresholdDate = (Get-Date).AddDays(-$Days)
    $LdapDate = $ThresholdDate.ToString("yyyyMMddHHmmss.0Z")

    # Define the Search Criteria with date filter
    $LdapFilter = "(&(whenCreated>=$LdapDate)(|(extensionAttribute6=*Non-Human Service Account*)(employeeID=*SVC*)(employeeNumber=*SVC*)))"

    # Query the Global Catalog (Port 3268)
    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server "${ForestRoot}:3268" -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled, Manager

    # Filter for accounts where Manager is NOT assigned (no limit)
    $OrphanedAccounts = $Users | Where-Object { [string]::IsNullOrEmpty($_.Manager) }

    $Results = foreach ($User in $OrphanedAccounts) {
        # Clean Domain Parsing
        $DnParts = $User.DistinguishedName.Split(',') | Where-Object { $_ -like "DC=*" }
        $CleanDomain = ($DnParts -replace "DC=", "") -join "."

        [PSCustomObject]@{
            Name              = $User.Name
            SamAccountName    = $User.SamAccountName
            AccountType       = "Service Account (Unmanaged)"
            Enabled           = $User.Enabled
            Domain            = $CleanDomain
            Created           = $User.whenCreated
            EA6_Value         = $User.extensionAttribute6
            EmployeeID        = $User.employeeID
            EmployeeNumber    = $User.employeeNumber
            ManagerStatus     = "MISSING"
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
