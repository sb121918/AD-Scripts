param(
    [string]$TargetDomain,
    [int]$Days = 90
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $serverParam = @{}
    if ($TargetDomain) {
        $serverParam.Server = $TargetDomain
    }

    $cutoffDate = (Get-Date).AddDays(-$Days)

    $filter = "Enabled -eq `$true -and LastLogonDate -ge '$cutoffDate' -and (SamAccountName -like 'svc*' -or SamAccountName -like 'sa-*' -or SamAccountName -like '*service*' -or Description -like '*service*')"

    $Users = Get-ADUser -Filter $filter @serverParam -ResultPageSize 2000 -ResultSetSize $null -Properties `
        DisplayName, SamAccountName, Enabled, LastLogonDate, LogonCount, `
        PasswordLastSet, PasswordNeverExpires, WhenCreated, Description, `
        Manager, AllowReversiblePasswordEncryption, DistinguishedName

    $Results = foreach ($User in $Users) {
        if ($User.LogonCount -gt 0) {
            [PSCustomObject]@{
                Name                 = $User.DisplayName
                SamAccountName       = $User.SamAccountName
                Enabled              = $User.Enabled
                LastLogon            = $User.LastLogonDate
                LogonCount           = $User.LogonCount
                PasswordLastSet      = $User.PasswordLastSet
                PasswordNeverExpires = $User.PasswordNeverExpires
                Created              = $User.WhenCreated
                Description          = $User.Description
                HasManager           = [bool]$User.Manager
                Domain               = ($User.DistinguishedName -split ',DC=' | Select-Object -Skip 1) -join '.'
                DistinguishedName    = $User.DistinguishedName
                RiskLevel            = if ($User.LogonCount -gt 100) { "High" } elseif ($User.LogonCount -gt 10) { "Medium" } else { "Low" }
            }
        }
    }

    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
