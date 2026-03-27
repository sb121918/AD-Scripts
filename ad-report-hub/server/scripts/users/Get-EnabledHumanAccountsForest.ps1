param(
    [string]$TargetDomain = "",
    [int]$Limit = 0
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

    # Enabled users + EA6 contains "Human"
    # Enabled is checked via userAccountControl bit (2 = disabled)
    # Exclude "Non-Human" values to avoid pulling service accounts tagged as Non-Human.
    $ldapFilter = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(extensionAttribute6=*Human*)(!(extensionAttribute6=*Non-Human*))(!(extensionAttribute6=*Non Human*)))"

    $results = New-Object System.Collections.Generic.List[PSObject]

    foreach ($domain in @($forest.Domains)) {
        try {
            $users = Get-ADUser -LDAPFilter $ldapFilter -Server $domain `
              -ResultPageSize 2000 -ResultSetSize $null `
              -Properties extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled, Manager

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

    @($results) | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}

