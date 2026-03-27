param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain,
    
    [Parameter(Mandatory=$false)]
    [int]$Days = 90
)

$ErrorActionPreference = 'SilentlyContinue'

function Format-DateForJSON {
    param($Date)
    if ($null -eq $Date -or $Date -eq [DateTime]::MinValue) { return $null }
    try {
        if ($Date -is [DateTime]) {
            return $Date.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        }
        $parsed = [DateTime]::Parse($Date)
        return $parsed.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    } catch { return $null }
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

try {
    $CutoffDate = (Get-Date).AddDays(-$Days)
    
    # Get service accounts that haven't logged on in X days
    # Service accounts typically have extensionAttribute6 = "SVC" or similar patterns
    $ServiceAccounts = Get-ADUser -Filter {
        Enabled -eq $true
    } -Server $TargetDomain -Properties `
        DisplayName, SamAccountName, Enabled, Created, Modified, LastLogonDate, `
        PasswordLastSet, PasswordNeverExpires, extensionAttribute6, Manager, `
        Department, Description, DistinguishedName `
        -ErrorAction Stop | Where-Object {
            # Filter for service accounts (EA6 contains SVC or account name patterns)
            ($_.extensionAttribute6 -like "*SVC*" -or 
             $_.SamAccountName -like "svc*" -or 
             $_.SamAccountName -like "*_svc" -or
             $_.SamAccountName -like "*-svc" -or
             $_.extensionAttribute6 -like "*Service*")
        } | Where-Object {
            # Filter for inactive (never logged on or not logged on in X days)
            ($null -eq $_.LastLogonDate) -or ($_.LastLogonDate -lt $CutoffDate)
        }

    if ($null -eq $ServiceAccounts -or @($ServiceAccounts).Count -eq 0) {
        @() | ConvertTo-Json
        exit 0
    }

    $Results = @($ServiceAccounts | ForEach-Object {
        $ManagerStatus = "NOT ASSIGNED"
        if ($_.Manager) {
            try {
                $ManagerObj = Get-ADUser -Identity $_.Manager -Server $TargetDomain -Properties DisplayName -ErrorAction SilentlyContinue
                if ($ManagerObj) {
                    $ManagerStatus = $ManagerObj.DisplayName
                } else {
                    $ManagerStatus = "MISSING"
                }
            } catch {
                $ManagerStatus = "MISSING"
            }
        }

        $DaysSinceLogon = if ($_.LastLogonDate) {
            [math]::Round((New-TimeSpan -Start $_.LastLogonDate -End (Get-Date)).TotalDays)
        } else {
            "Never"
        }

        [PSCustomObject]@{
            Name = $_.DisplayName
            SamAccountName = $_.SamAccountName
            AccountType = "Service Account"
            Enabled = $_.Enabled
            Domain = $TargetDomain
            Created = Format-DateForJSON $_.Created
            LastLogonDate = Format-DateForJSON $_.LastLogonDate
            DaysSinceLogon = $DaysSinceLogon
            PasswordLastSet = Format-DateForJSON $_.PasswordLastSet
            PasswordNeverExpires = $_.PasswordNeverExpires
            EA6_Value = $_.extensionAttribute6
            ManagerStatus = $ManagerStatus
            Department = $_.Department
            Description = $_.Description
            DistinguishedName = $_.DistinguishedName
        }
    } | Sort-Object DaysSinceLogon -Descending)

    @($Results) | ConvertTo-Json -Depth 4
}
catch {
    @{ Error = "Failed to query inactive service accounts: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}
