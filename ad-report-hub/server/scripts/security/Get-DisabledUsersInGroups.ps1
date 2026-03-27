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
            $DisabledUsers = Get-ADUser -Filter 'Enabled -eq $false' -Server $domain -Properties `
                DisplayName, SamAccountName, mail, Title, Department, Enabled, `
                memberOf, WhenCreated, WhenChanged, DistinguishedName -ErrorAction SilentlyContinue

            foreach ($User in @($DisabledUsers)) {
                if ($User.memberOf -and @($User.memberOf).Count -gt 0) {
                    $groups = @($User.memberOf)

                    $privilegedGroups = $groups | Where-Object {
                        $_ -match 'Domain Admins|Enterprise Admins|Schema Admins|Administrators|Account Operators|Backup Operators'
                    }

                    $Results.Add([PSCustomObject]@{
                        Name              = $User.DisplayName
                        SamAccountName    = $User.SamAccountName
                        Email             = $User.mail
                        Title             = $User.Title
                        Department        = $User.Department
                        Enabled           = $User.Enabled
                        GroupCount        = $groups.Count
                        PrivilegedGroups  = @($privilegedGroups).Count
                        Created           = $User.WhenCreated
                        Modified          = $User.WhenChanged
                        Domain            = $domain
                        DistinguishedName = $User.DistinguishedName
                        RiskLevel         = if (@($privilegedGroups).Count -gt 0) { "Critical" } elseif ($groups.Count -gt 10) { "High" } else { "Medium" }
                    })
                }
            }
        } catch { }
    }

    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
