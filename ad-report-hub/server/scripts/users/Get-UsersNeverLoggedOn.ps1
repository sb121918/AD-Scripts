param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all users who never logged on (no limit)
    $Users = Get-ADUser -Filter { LastLogonDate -notlike "*" -and Enabled -eq $true } -Properties DisplayName, EmailAddress, Department, Title, WhenCreated
    
    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.EmailAddress
            Department        = $User.Department
            Title             = $User.Title
            Enabled           = $User.Enabled
            Created           = $User.WhenCreated
            LastLogon         = "Never"
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
