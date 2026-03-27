param(
    [int]$Days = 60
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $CutoffDate = (Get-Date).AddDays(-$Days)
    
    # Fetch all users not logged on in X days (no limit)
    $Users = Get-ADUser -Filter { LastLogonDate -lt $CutoffDate -and Enabled -eq $true } -Properties DisplayName, EmailAddress, Department, Title, LastLogonDate, WhenCreated |
             Sort-Object LastLogonDate
    
    $Results = foreach ($User in $Users) {
        $DaysSinceLogon = if ($User.LastLogonDate) { ((Get-Date) - $User.LastLogonDate).Days } else { "N/A" }
        
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.EmailAddress
            Department        = $User.Department
            Title             = $User.Title
            Enabled           = $User.Enabled
            Created           = $User.WhenCreated
            LastLogon         = $User.LastLogonDate
            DaysSinceLogon    = $DaysSinceLogon
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
