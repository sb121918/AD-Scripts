param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $WarningPreference = 'SilentlyContinue'

    $anchorDomain = if ($TargetDomain) { $TargetDomain } else { (Get-ADDomain).DNSRoot }
    $forest = Get-ADForest -Server $anchorDomain -ErrorAction Stop
    $domainsToQuery = @($forest.Domains)

    $Results = New-Object System.Collections.Generic.List[PSObject]

    foreach ($domain in $domainsToQuery) {
        try {
            $Users = Get-ADUser -Filter 'adminCount -eq 1' -Server $domain -Properties `
                DisplayName, SamAccountName, mail, Title, Department, Enabled, `
                memberOf, WhenCreated, LastLogonDate, PasswordLastSet, adminCount, `
                DistinguishedName -ErrorAction SilentlyContinue

            foreach ($User in @($Users)) {
                $GroupCount = if ($User.memberOf) { @($User.memberOf).Count } else { 0 }

                $Results.Add([PSCustomObject]@{
                    Name              = $User.DisplayName
                    SamAccountName    = $User.SamAccountName
                    Email             = $User.mail
                    Title             = $User.Title
                    Department        = $User.Department
                    Enabled           = $User.Enabled
                    GroupCount        = $GroupCount
                    Created           = $User.WhenCreated
                    LastLogon         = $User.LastLogonDate
                    PasswordLastSet   = $User.PasswordLastSet
                    AdminCount        = $User.adminCount
                    Domain            = $domain
                    DistinguishedName = $User.DistinguishedName
                    RiskLevel         = if (-not $User.Enabled) { "Medium" } elseif (-not $User.LastLogonDate -or $User.LastLogonDate -lt (Get-Date).AddDays(-90)) { "High" } else { "Standard" }
                })
            }
        } catch { }
    }

    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
