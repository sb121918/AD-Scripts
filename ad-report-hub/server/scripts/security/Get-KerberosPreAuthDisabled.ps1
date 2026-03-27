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
            $Users = Get-ADUser -Filter 'DoesNotRequirePreAuth -eq $true' -Server $domain -Properties `
                DisplayName, SamAccountName, mail, Title, Department, Enabled, `
                DoesNotRequirePreAuth, WhenCreated, LastLogonDate, PasswordLastSet, `
                DistinguishedName -ErrorAction SilentlyContinue

            foreach ($User in @($Users)) {
                $Results.Add([PSCustomObject]@{
                    Name                = $User.DisplayName
                    SamAccountName      = $User.SamAccountName
                    Email               = $User.mail
                    Title               = $User.Title
                    Department          = $User.Department
                    Enabled             = $User.Enabled
                    PreAuthDisabled     = $User.DoesNotRequirePreAuth
                    Created             = $User.WhenCreated
                    LastLogon           = $User.LastLogonDate
                    PasswordLastSet     = $User.PasswordLastSet
                    Domain              = $domain
                    DistinguishedName   = $User.DistinguishedName
                    RiskLevel           = "High"
                    Vulnerability       = "AS-REP Roasting"
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
