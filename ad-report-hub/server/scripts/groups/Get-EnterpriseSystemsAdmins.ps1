param(
    [string]$ForestName = "",
    [string]$ForestDomain = ""
)

$ErrorActionPreference = 'Stop'

# Helper function to format dates as ISO 8601 strings for JavaScript
function Format-DateForJSON {
    param($Date)
    if ($null -eq $Date) { return $null }
    try {
        return $Date.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    } catch {
        return $null
    }
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$Results = @()

try {
    if (-not $ForestName -or ($ForestName.ToLower() -ne "viacom")) {
        # This group is Viacom-only. If the selected forest isn't Viacom, return empty.
        @() | ConvertTo-Json
        exit 0
    }

    $forest = $null
    try {
        if ($ForestDomain) {
            $forest = Get-ADForest -Server $ForestDomain -ErrorAction Stop
        } else {
            $forest = Get-ADForest -ErrorAction Stop
        }
    } catch {
        $forest = $null
    }

    $domainsToSearch = @()
    if ($forest -and $forest.Domains) {
        $domainsToSearch = @($forest.Domains)
    } elseif ($ForestDomain) {
        $domainsToSearch = @($ForestDomain)
    }
    if (-not $domainsToSearch -or $domainsToSearch.Count -eq 0) {
        @() | ConvertTo-Json
        exit 0
    }
    
    # Try multiple possible group names
    $GroupNames = @(
        "AD-Enterprise Systems Admins",
        "AD-Enterprise-Systems Admins", 
        "AD-Enterprise-System Admin",
        "AD-Enterprise System Admins",
        "Enterprise Systems Admins"
    )
    
    $Group = $null
    $FoundGroupName = $null
    $FoundDomain = $null
    $ServerToUse = $null
    
    foreach ($d in $domainsToSearch) {
        # Prefer PDC emulator for consistent reads
        $dc = $null
        try {
            $domainInfo = Get-ADDomain -Server $d -ErrorAction Stop
            $dc = $domainInfo.PDCEmulator
        } catch {
            try { $dc = (Get-ADDomainController -DomainName $d -Discover -ErrorAction Stop).HostName } catch { $dc = $d }
        }

        foreach ($GroupName in $GroupNames) {
            try {
                $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -Server $dc -Properties WhenCreated, WhenChanged, Description -ErrorAction SilentlyContinue
                if ($Group) {
                    $FoundGroupName = $Group.Name
                    $FoundDomain = $d
                    $ServerToUse = $dc
                    break
                }
            } catch { }
        }

        if (-not $Group) {
            try {
                $Group = Get-ADGroup -Filter "Name -like '*Enterprise*System*Admin*'" -Server $dc -Properties WhenCreated, WhenChanged, Description -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($Group) {
                    $FoundGroupName = $Group.Name
                    $FoundDomain = $d
                    $ServerToUse = $dc
                }
            } catch { }
        }

        if ($Group) { break }
    }
    
    if (-not $Group) {
        # Return empty array instead of error - group might not exist
        @() | ConvertTo-Json
        exit 0
    }
    
    # Get members - don't use -Recursive to avoid referral issues
    $Members = @()
    try {
        $Members = Get-ADGroupMember -Identity $Group.DistinguishedName -Server $ServerToUse -ErrorAction Stop | 
                   Where-Object { $_.objectClass -eq 'user' }
    } catch {
        # If direct member query fails, try using the group's DN
        try {
            $GroupDN = $Group.DistinguishedName
            $Members = Get-ADUser -Filter "memberOf -eq '$GroupDN'" -Server $ServerToUse -Properties DisplayName, EmailAddress, Title, Department, Enabled, WhenCreated -ErrorAction SilentlyContinue
        } catch { }
    }
    
    foreach ($Member in $Members) {
        try {
            $User = $null
            
            # If member already has user properties, use them
            if ($Member.SamAccountName -and $Member.Enabled -ne $null) {
                $User = $Member
            } else {
                # Get full user details
                $User = Get-ADUser -Identity $Member.distinguishedName -Server $ServerToUse -Properties DisplayName, EmailAddress, Title, Department, Enabled, WhenCreated -ErrorAction SilentlyContinue
            }
            
            if ($User) {
                $Results += [PSCustomObject]@{
                    GroupName         = $FoundGroupName
                    Domain            = $FoundDomain
                    MemberName        = $User.Name
                    SamAccountName    = $User.SamAccountName
                    DisplayName       = $User.DisplayName
                    Email             = $User.EmailAddress
                    Title             = $User.Title
                    Department        = $User.Department
                    Enabled           = $User.Enabled
                    MemberCreated     = Format-DateForJSON $User.WhenCreated
                    GroupModified     = Format-DateForJSON $Group.WhenChanged
                    DistinguishedName = $User.DistinguishedName
                }
            }
        } catch { }
    }
}
catch {
    @{ Error = "Failed to query AD-Enterprise Systems Admins: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($Results) | ConvertTo-Json -Depth 3
