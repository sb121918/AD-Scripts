param(
    [string]$Filter = "all",
    [string]$Format = "json",
    [int]$Limit = 1000
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Build filter based on parameter
    $ADFilter = switch ($Filter) {
        "enabled"  { { Enabled -eq $true } }
        "disabled" { { Enabled -eq $false } }
        default    { "*" }
    }
    
    $Users = Get-ADUser -Filter $ADFilter -Properties `
        DisplayName, EmailAddress, employeeID, employeeNumber, Title, Department, `
        telephoneNumber, mobile, Manager, Enabled, WhenCreated, WhenChanged, LastLogonDate |
        Select-Object -First $Limit
    
    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            UserPrincipalName = $User.UserPrincipalName
            Email             = $User.EmailAddress
            EmployeeID        = $User.employeeID
            EmployeeNumber    = $User.employeeNumber
            Title             = $User.Title
            Department        = $User.Department
            Telephone         = $User.telephoneNumber
            Mobile            = $User.mobile
            Manager           = if ($User.Manager) { ($User.Manager -split ',')[0].Replace("CN=","") } else { $null }
            Enabled           = $User.Enabled
            Created           = $User.WhenCreated
            Modified          = $User.WhenChanged
            LastLogon         = $User.LastLogonDate
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
