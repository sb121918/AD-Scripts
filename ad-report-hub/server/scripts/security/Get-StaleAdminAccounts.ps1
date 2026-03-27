param(
    [string]$TargetDomain,
    [int]$InactiveDays = 90
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $WarningPreference = 'SilentlyContinue'

    $anchorDomain = if ($TargetDomain) { $TargetDomain } else { (Get-ADDomain).DNSRoot }
    $forest = Get-ADForest -Server $anchorDomain -ErrorAction Stop
    $domainsToQuery = @($forest.Domains)
    $cutoffDate = (Get-Date).AddDays(-$InactiveDays)

    $Results = New-Object System.Collections.Generic.List[PSObject]

    foreach ($domain in $domainsToQuery) {
        try {
            $adminUsers = Get-ADUser -Filter 'adminCount -eq 1' -Server $domain -Properties `
                DisplayName, SamAccountName, mail, Title, Department, Enabled, `
                memberOf, WhenCreated, LastLogonDate, PasswordLastSet, PasswordNeverExpires, `
                DistinguishedName -ErrorAction SilentlyContinue

            foreach ($User in @($adminUsers)) {
                $isStale = $false
                $staleReason = @()

                if (-not $User.LastLogonDate) {
                    $isStale = $true
                    $staleReason += "Never logged on"
                }
                elseif ($User.LastLogonDate -lt $cutoffDate) {
                    $isStale = $true
                    $staleReason += "Inactive for $InactiveDays+ days"
                }

                if (-not $User.Enabled) {
                    $isStale = $true
                    $staleReason += "Account disabled"
                }

                if ($isStale) {
                    $daysSinceLogon = if ($User.LastLogonDate) {
                        [math]::Round(((Get-Date) - $User.LastLogonDate).TotalDays)
                    } else { "Never" }

                    $groupCount = if ($User.memberOf) { @($User.memberOf).Count } else { 0 }

                    $Results.Add([PSCustomObject]@{
                        Name              = $User.DisplayName
                        SamAccountName    = $User.SamAccountName
                        Email             = $User.mail
                        Title             = $User.Title
                        Department        = $User.Department
                        Enabled           = $User.Enabled
                        GroupCount        = $groupCount
                        Created           = $User.WhenCreated
                        LastLogon         = $User.LastLogonDate
                        DaysSinceLogon    = $daysSinceLogon
                        PasswordLastSet   = $User.PasswordLastSet
                        PasswordNeverExpires = $User.PasswordNeverExpires
                        StaleReason       = $staleReason -join "; "
                        Domain            = $domain
                        DistinguishedName = $User.DistinguishedName
                        RiskLevel         = if (-not $User.Enabled -and $groupCount -gt 0) { "Critical" } elseif ($daysSinceLogon -eq "Never") { "High" } else { "Medium" }
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
