param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $serverParam = @{}
    if ($TargetDomain) {
        $serverParam.Server = $TargetDomain
    }

    $filter = "Enabled -eq `$true -and PasswordNeverExpires -eq `$true -and (SamAccountName -like 'svc*' -or SamAccountName -like 'sa-*' -or SamAccountName -like '*service*' -or Description -like '*service*')"

    $Users = Get-ADUser -Filter $filter @serverParam -ResultPageSize 2000 -ResultSetSize $null -Properties `
        DisplayName, SamAccountName, Enabled, PasswordNeverExpires, PasswordLastSet, `
        LastLogonDate, WhenCreated, Description, Manager, DistinguishedName

    $Results = foreach ($User in $Users) {
        $daysSincePwdSet = if ($User.PasswordLastSet) {
            [math]::Round(((Get-Date) - $User.PasswordLastSet).TotalDays)
        } else { "Never" }

        [PSCustomObject]@{
            Name                 = $User.DisplayName
            SamAccountName       = $User.SamAccountName
            Enabled              = $User.Enabled
            PasswordNeverExpires = $User.PasswordNeverExpires
            PasswordLastSet      = $User.PasswordLastSet
            DaysSincePasswordSet = $daysSincePwdSet
            LastLogon            = $User.LastLogonDate
            Created              = $User.WhenCreated
            Description          = $User.Description
            HasManager           = [bool]$User.Manager
            Domain               = ($User.DistinguishedName -split ',DC=' | Select-Object -Skip 1) -join '.'
            DistinguishedName    = $User.DistinguishedName
            RiskLevel            = if ($daysSincePwdSet -eq "Never" -or $daysSincePwdSet -gt 365) { "High" } elseif ($daysSincePwdSet -gt 180) { "Medium" } else { "Low" }
        }
    }

    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
