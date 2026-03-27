param(
    [string]$TargetDomain = "",
    [int]$Limit = 0,
    [string]$OutputPath = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $WarningPreference = 'SilentlyContinue'
    $ErrorActionPreference = 'SilentlyContinue'

    if (-not $TargetDomain) {
        @{ Error = "TargetDomain is required to discover forest domains." } | ConvertTo-Json
        exit 1
    }

    function Get-CnFromDn {
        param([string]$Dn)
        if (-not $Dn) { return $null }
        if ($Dn -match '^CN=([^,]+),') { return $matches[1] }
        return $Dn
    }

    $forest = $null
    try {
        $forest = Get-ADForest -Server $TargetDomain -ErrorAction Stop
    } catch {
        $forest = $null
    }

    if (-not $forest -or -not $forest.Domains) {
        @{ Error = "Failed to discover forest domains from TargetDomain: $TargetDomain" } | ConvertTo-Json
        exit 1
    }

    # All users (enabled + disabled) where EA6 contains "Human"
    # Exclude "Non-Human" values to avoid pulling service accounts tagged as Non-Human.
    $ldapFilter = "(&(objectCategory=person)(objectClass=user)(extensionAttribute6=*Human*)(!(extensionAttribute6=*Non-Human*))(!(extensionAttribute6=*Non Human*)))"

    $results = New-Object System.Collections.Generic.List[PSObject]

    $domainCount = @($forest.Domains).Count
    $currentDomain = 0

    foreach ($domain in @($forest.Domains)) {
        $currentDomain++
        Write-Host "[$currentDomain/$domainCount] Querying domain: $domain ..." -ForegroundColor Cyan
        
        try {
            $users = Get-ADUser -LDAPFilter $ldapFilter -Server $domain `
              -ResultPageSize 2000 -ResultSetSize $null `
              -Properties extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled, Manager

            $domainUserCount = @($users).Count
            Write-Host "  Found $domainUserCount users in $domain" -ForegroundColor Green

            foreach ($u in @($users)) {
                $results.Add([PSCustomObject]@{
                    Name              = $u.Name
                    SamAccountName    = $u.SamAccountName
                    Enabled           = $u.Enabled
                    EA6_Value         = $u.extensionAttribute6
                    EmployeeID        = $u.employeeID
                    EmployeeNumber    = $u.employeeNumber
                    Title             = $u.Title
                    Department        = $u.Department
                    Description       = $u.Description
                    ManagerName       = Get-CnFromDn $u.Manager
                    Created           = if ($u.whenCreated) { $u.whenCreated.ToString('yyyy-MM-dd') } else { $null }
                    Domain            = $domain
                    DistinguishedName = $u.DistinguishedName
                })

                if ($Limit -gt 0 -and $results.Count -ge $Limit) { break }
            }
        } catch { }

        if ($Limit -gt 0 -and $results.Count -ge $Limit) { break }
    }

    Write-Host "`nExporting $($results.Count) users to CSV..." -ForegroundColor Yellow

    if (-not $OutputPath) {
        $OutputPath = "HumanAccounts_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    }

    @($results) | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "Export complete: $OutputPath" -ForegroundColor Green

    @{ 
        Success = $true
        TotalCount = $results.Count
        OutputFile = (Resolve-Path $OutputPath).Path
    } | ConvertTo-Json
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
