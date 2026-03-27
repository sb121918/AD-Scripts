param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all disabled users (no limit)
    $Users = Get-ADUser -Filter { Enabled -eq $false } -Properties DisplayName, EmailAddress, Department, Title, WhenCreated, WhenChanged
    
    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.EmailAddress
            Department        = $User.Department
            Title             = $User.Title
            Enabled           = $false
            Created           = $User.WhenCreated
            Modified          = $User.WhenChanged
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
