param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Fetch all users with expired passwords (no limit)
    $Users = Search-ADAccount -PasswordExpired -UsersOnly |
             Get-ADUser -Properties DisplayName, EmailAddress, Department, Title, PasswordLastSet, PasswordExpired
    
    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.EmailAddress
            Department        = $User.Department
            Title             = $User.Title
            PasswordExpired   = $User.PasswordExpired
            PasswordLastSet   = $User.PasswordLastSet
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
