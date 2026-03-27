param(
    [int]$Days = 30
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $CutoffDate = (Get-Date).AddDays(-$Days)
    
    # Fetch all users created in last X days (no limit)
    $Users = Get-ADUser -Filter { WhenCreated -ge $CutoffDate } -Properties DisplayName, EmailAddress, Department, Title, Enabled, WhenCreated |
             Sort-Object WhenCreated -Descending
    
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
