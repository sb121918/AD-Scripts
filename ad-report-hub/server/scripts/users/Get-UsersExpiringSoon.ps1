param(
    [int]$DaysUntilExpiry = 30,
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $serverParam = @{}
    if ($TargetDomain) {
        $serverParam.Server = $TargetDomain
    }
    
    $today = Get-Date
    $expiryThreshold = $today.AddDays($DaysUntilExpiry)
    
    # Find users with AccountExpirationDate set and expiring within threshold
    $Users = Get-ADUser -Filter 'AccountExpirationDate -like "*"' @serverParam -Properties `
        DisplayName, SamAccountName, mail, Title, Department, Enabled, `
        AccountExpirationDate, WhenCreated, LastLogonDate, Manager, `
        EmployeeType, DistinguishedName
    
    $Results = foreach ($User in $Users) {
        if ($User.AccountExpirationDate -and $User.AccountExpirationDate -le $expiryThreshold) {
            $daysRemaining = [math]::Round(($User.AccountExpirationDate - $today).TotalDays)
            $status = if ($daysRemaining -lt 0) { "Expired" } elseif ($daysRemaining -eq 0) { "Expires Today" } elseif ($daysRemaining -le 7) { "Critical" } elseif ($daysRemaining -le 14) { "Warning" } else { "Upcoming" }
            
            $managerName = $null
            if ($User.Manager) {
                try {
                    $mgr = Get-ADUser -Identity $User.Manager @serverParam -Properties DisplayName -ErrorAction SilentlyContinue
                    $managerName = $mgr.DisplayName
                } catch {}
            }
            
            [PSCustomObject]@{
                Name              = $User.DisplayName
                SamAccountName    = $User.SamAccountName
                Email             = $User.mail
                Title             = $User.Title
                Department        = $User.Department
                EmployeeType      = $User.EmployeeType
                Enabled           = $User.Enabled
                ExpirationDate    = $User.AccountExpirationDate
                DaysRemaining     = $daysRemaining
                ExpiryStatus      = $status
                Created           = $User.WhenCreated
                LastLogon         = $User.LastLogonDate
                Manager           = $managerName
                Domain            = ($User.DistinguishedName -split ',DC=' | Select-Object -Skip 1) -join '.'
                DistinguishedName = $User.DistinguishedName
            }
        }
    }
    
    @($Results) | Sort-Object DaysRemaining | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
