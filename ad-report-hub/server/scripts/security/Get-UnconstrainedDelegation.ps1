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
            # Users with unconstrained delegation
            $Users = Get-ADUser -Filter 'TrustedForDelegation -eq $true' -Server $domain -Properties `
                DisplayName, SamAccountName, mail, Enabled, TrustedForDelegation, `
                servicePrincipalName, WhenCreated, LastLogonDate, DistinguishedName -ErrorAction SilentlyContinue

            foreach ($User in @($Users)) {
                $Results.Add([PSCustomObject]@{
                    ObjectType        = "User"
                    Name              = $User.DisplayName
                    SamAccountName    = $User.SamAccountName
                    Email             = $User.mail
                    Enabled           = $User.Enabled
                    DelegationType    = "Unconstrained"
                    SPNCount          = if ($User.servicePrincipalName) { @($User.servicePrincipalName).Count } else { 0 }
                    Created           = $User.WhenCreated
                    LastLogon         = $User.LastLogonDate
                    Domain            = $domain
                    DistinguishedName = $User.DistinguishedName
                    RiskLevel         = "Critical"
                })
            }

            # Computers with unconstrained delegation (excluding DCs)
            $Computers = Get-ADComputer -Filter 'TrustedForDelegation -eq $true' -Server $domain -Properties `
                Name, DNSHostName, Enabled, TrustedForDelegation, OperatingSystem, `
                servicePrincipalName, WhenCreated, LastLogonDate, DistinguishedName, PrimaryGroupID -ErrorAction SilentlyContinue

            foreach ($Computer in @($Computers)) {
                if ($Computer.PrimaryGroupID -in @(516, 521)) { continue }

                $Results.Add([PSCustomObject]@{
                    ObjectType        = "Computer"
                    Name              = $Computer.Name
                    SamAccountName    = $Computer.Name + '$'
                    Email             = $null
                    Enabled           = $Computer.Enabled
                    DelegationType    = "Unconstrained"
                    SPNCount          = if ($Computer.servicePrincipalName) { @($Computer.servicePrincipalName).Count } else { 0 }
                    Created           = $Computer.WhenCreated
                    LastLogon         = $Computer.LastLogonDate
                    Domain            = $domain
                    DistinguishedName = $Computer.DistinguishedName
                    RiskLevel         = "Critical"
                    OperatingSystem   = $Computer.OperatingSystem
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
