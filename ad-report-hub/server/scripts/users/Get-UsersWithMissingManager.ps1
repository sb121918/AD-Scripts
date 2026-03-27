param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $serverParam = @{}
    if ($TargetDomain) {
        $serverParam.Server = $TargetDomain
    }
    
    # Find enabled users with no manager set
    $Users = Get-ADUser -Filter 'Enabled -eq $true -and Manager -notlike "*"' @serverParam -Properties `
        DisplayName, SamAccountName, mail, Title, Department, Enabled, `
        WhenCreated, LastLogonDate, EmployeeType, extensionAttribute6, `
        DistinguishedName
    
    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            Name              = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.mail
            Title             = $User.Title
            Department        = $User.Department
            EmployeeType      = $User.EmployeeType
            EA6               = $User.extensionAttribute6
            Enabled           = $User.Enabled
            Created           = $User.WhenCreated
            LastLogon         = $User.LastLogonDate
            Manager           = $null
            ManagerStatus     = "Missing"
            Domain            = ($User.DistinguishedName -split ',DC=' | Select-Object -Skip 1) -join '.'
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
