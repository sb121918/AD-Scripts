param(
    [int]$Days = 30,
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $serverParam = @{}
    if ($TargetDomain) {
        $serverParam.Server = $TargetDomain
    }
    
    $cutoffDate = (Get-Date).AddDays(-$Days)
    
    $Users = Get-ADUser -Filter "WhenChanged -ge '$cutoffDate'" @serverParam -Properties `
        DisplayName, SamAccountName, mail, Title, Department, Enabled, `
        WhenCreated, WhenChanged, LastLogonDate, PasswordLastSet, `
        Manager, DistinguishedName
    
    $Results = foreach ($User in $Users) {
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
            Enabled           = $User.Enabled
            Created           = $User.WhenCreated
            Modified          = $User.WhenChanged
            LastLogon         = $User.LastLogonDate
            PasswordLastSet   = $User.PasswordLastSet
            Manager           = $managerName
            Domain            = ($User.DistinguishedName -split ',DC=' | Select-Object -Skip 1) -join '.'
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | Sort-Object Modified -Descending | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
