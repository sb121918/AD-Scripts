param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $Users = Get-ADUser -Filter 'Enabled -eq $true' -Server $TargetDomain -ResultPageSize 2000 -ResultSetSize $null -Properties `
        DisplayName, EmailAddress, Title, Department, Company, LastLogonDate, whenCreated

    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            Name              = $User.Name
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            UPN               = $User.UserPrincipalName
            Email             = $User.EmailAddress
            Title             = $User.Title
            Department        = $User.Department
            Company           = $User.Company
            LastLogon         = $User.LastLogonDate
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
