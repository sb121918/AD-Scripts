# Get-OUTree.ps1
# Returns hierarchical OU data with Children arrays, child object counts, and linked GPOs
param(
    [string]$Domain = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    if ($Domain) {
        $dc = (Get-ADDomainController -DomainName $Domain -Discover -ErrorAction Stop).HostName[0]
        $domainDN = (Get-ADDomain -Server $dc).DistinguishedName
    } else {
        $dc = $null
        $domainDN = (Get-ADDomain).DistinguishedName
    }

    $getADParams = @{ ErrorAction = 'SilentlyContinue' }
    if ($dc) { $getADParams.Server = $dc }

    # Get all OUs
    $allOUs = Get-ADOrganizationalUnit -Filter * -Properties Name, Description, DistinguishedName, gPLink, ProtectedFromAccidentalDeletion @getADParams

    # Build a lookup table by DN
    $ouLookup = @{}
    foreach ($ou in $allOUs) {
        $ouLookup[$ou.DistinguishedName] = @{
            Name = $ou.Name
            Description = if ($ou.Description) { $ou.Description } else { '' }
            DistinguishedName = $ou.DistinguishedName
            Protected = [bool]$ou.ProtectedFromAccidentalDeletion
            GPOLinks = @()
            Children = @()
            UserCount = 0
            GroupCount = 0
            ComputerCount = 0
        }

        # Parse GPO links
        if ($ou.gPLink) {
            $gpoPattern = '\[LDAP://cn=\{([0-9a-fA-F\-]+)\},[^\]]+\]'
            $matches = [regex]::Matches($ou.gPLink, $gpoPattern)
            foreach ($match in $matches) {
                $gpoGuid = $match.Groups[1].Value
                try {
                    $gpo = Get-ADObject -Filter "objectClass -eq 'groupPolicyContainer' -and Name -eq '{$gpoGuid}'" -Properties displayName @getADParams
                    if ($gpo) {
                        $ouLookup[$ou.DistinguishedName].GPOLinks += $gpo.displayName
                    } else {
                        $ouLookup[$ou.DistinguishedName].GPOLinks += "{$gpoGuid}"
                    }
                } catch {
                    $ouLookup[$ou.DistinguishedName].GPOLinks += "{$gpoGuid}"
                }
            }
        }
    }

    # Count child objects per OU (users, groups, computers)
    foreach ($ou in $allOUs) {
        $dn = $ou.DistinguishedName
        try {
            $ouLookup[$dn].UserCount = @(Get-ADUser -SearchBase $dn -SearchScope OneLevel -Filter * @getADParams).Count
            $ouLookup[$dn].GroupCount = @(Get-ADGroup -SearchBase $dn -SearchScope OneLevel -Filter * @getADParams).Count
            $ouLookup[$dn].ComputerCount = @(Get-ADComputer -SearchBase $dn -SearchScope OneLevel -Filter * @getADParams).Count
        } catch {
            # Skip count errors
        }
    }

    # Build tree structure
    foreach ($ou in $allOUs) {
        $dn = $ou.DistinguishedName
        # Find parent OU DN
        $parts = $dn -split ',', 2
        if ($parts.Count -gt 1) {
            $parentDN = $parts[1]
            if ($ouLookup.ContainsKey($parentDN)) {
                $ouLookup[$parentDN].Children += $ouLookup[$dn]
            }
        }
    }

    # Find root OUs (those whose parent is the domain DN, not another OU)
    $rootOUs = @()
    foreach ($ou in $allOUs) {
        $dn = $ou.DistinguishedName
        $parts = $dn -split ',', 2
        if ($parts.Count -gt 1) {
            $parentDN = $parts[1]
            if (-not $ouLookup.ContainsKey($parentDN)) {
                $rootOUs += $ouLookup[$dn]
            }
        }
    }

    # Build the final tree
    $tree = @{
        Name = $domainDN
        DistinguishedName = $domainDN
        Description = "Domain Root"
        Protected = $false
        GPOLinks = @()
        Children = $rootOUs
        UserCount = 0
        GroupCount = 0
        ComputerCount = 0
    }

    $tree | ConvertTo-Json -Depth 15
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
}
