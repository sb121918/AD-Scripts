param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all enabled users (no limit)
    $Users = Get-ADUser -Filter { Enabled -eq $true } -Properties DisplayName, EmailAddress, Department, Title, WhenCreated, LastLogonDate
    
    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.EmailAddress
            Department        = $User.Department
            Title             = $User.Title
            Enabled           = $true
            Created           = $User.WhenCreated
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
