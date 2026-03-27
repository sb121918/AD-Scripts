param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all users (no limit)
    $Users = Get-ADUser -Filter * -Properties DisplayName, EmailAddress, Department, Title, Enabled, WhenCreated
    
    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.EmailAddress
            Department        = $User.Department
            Title             = $User.Title
            Enabled           = $User.Enabled
            Created           = $User.WhenCreated
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
