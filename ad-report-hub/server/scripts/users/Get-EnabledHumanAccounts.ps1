param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $LdapFilter = "(extensionAttribute6=Human Primary Identity SF Match)"

    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server $TargetDomain -ResultPageSize 2000 -ResultSetSize $null -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled

    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            Name              = $User.Name
            SamAccountName    = $User.SamAccountName
            Enabled           = $User.Enabled
            EA6_Value         = $User.extensionAttribute6
            EmployeeID        = $User.employeeID
            EmployeeNumber    = $User.employeeNumber
            Title             = $User.Title
            Department        = $User.Department
            Description       = $User.Description
            Created           = $User.whenCreated
            Domain            = $TargetDomain
            DistinguishedName = $User.DistinguishedName
        }
    }

    @($Results) | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
