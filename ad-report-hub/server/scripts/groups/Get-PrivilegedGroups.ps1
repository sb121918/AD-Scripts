param(
    [string]$ForestDomain = "",
    [string]$ForestName = "",
    [string]$PrivilegedGroupsCsv = ""
)

$ErrorActionPreference = 'SilentlyContinue'

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

# Helper function to get group members safely
function Get-SafeGroupMembers {
    param(
        [Parameter(Mandatory=$true)]
        $Group,
        [Parameter(Mandatory=$true)]
        [string]$Server
    )
    
    $Members = @()
    
    # Try Get-ADGroupMember first (without -Recursive to avoid referral issues)
    try {
        $Members = Get-ADGroupMember -Identity $Group.DistinguishedName -Server $Server -ErrorAction Stop | 
                   Where-Object { $_.objectClass -eq 'user' }
    } catch {
        # If that fails, try querying users directly by memberOf
        try {
            $GroupDN = $Group.DistinguishedName
            $Members = Get-ADUser -Filter "memberOf -eq '$GroupDN'" -Server $Server -ErrorAction SilentlyContinue
        } catch { }
    }
    
    return $Members
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$AllPrivilegedUsers = @()

try {
    $Forest = $null
    try {
        if ($ForestDomain) {
            $Forest = Get-ADForest -Server $ForestDomain -ErrorAction Stop
        } else {
            $Forest = Get-ADForest -ErrorAction Stop
        }
    } catch {
        $Forest = $null
    }

    if (-not $Forest) {
        @{ Error = "Failed to query forest (Get-ADForest). Provide a valid ForestDomain." } | ConvertTo-Json
        exit 1
    }
    $RootDomain = $Forest.RootDomain

    # Optional: allow server to provide the exact privileged groups list
    $PrivilegedGroupsFromServer = @()
    if ($PrivilegedGroupsCsv) {
        $PrivilegedGroupsFromServer = $PrivilegedGroupsCsv.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }

    # 1. Get Enterprise Admins from root domain
    try {
        $EAGroup = Get-ADGroup -Identity "Enterprise Admins" -Server $RootDomain -Properties WhenChanged -ErrorAction Stop
        $EAMembers = Get-SafeGroupMembers -Group $EAGroup -Server $RootDomain
        
        foreach ($Member in $EAMembers) {
            try {
                $User = Get-ADUser -Identity $Member.distinguishedName -Server $RootDomain -Properties DisplayName, EmailAddress, Title, Department, Enabled, WhenCreated -ErrorAction SilentlyContinue
                
                if ($User) {
                    $AllPrivilegedUsers += [PSCustomObject]@{
                        GroupName         = "Enterprise Admins"
                        Domain            = $RootDomain
                        MemberName        = $User.Name
                        SamAccountName    = $User.SamAccountName
                        DisplayName       = $User.DisplayName
                        Email             = $User.EmailAddress
                        Title             = $User.Title
                        Department        = $User.Department
                        Enabled           = $User.Enabled
                        MemberCreated     = Format-DateForJSON $User.WhenCreated
                        GroupModified     = Format-DateForJSON $EAGroup.WhenChanged
                        DistinguishedName = $User.DistinguishedName
                    }
                }
            } catch { }
        }
    } catch { }

    # 2. Get Schema Admins from root domain
    try {
        $SAGroup = Get-ADGroup -Identity "Schema Admins" -Server $RootDomain -Properties WhenChanged -ErrorAction Stop
        $SAMembers = Get-SafeGroupMembers -Group $SAGroup -Server $RootDomain
        
        foreach ($Member in $SAMembers) {
            try {
                $User = Get-ADUser -Identity $Member.distinguishedName -Server $RootDomain -Properties DisplayName, EmailAddress, Title, Department, Enabled, WhenCreated -ErrorAction SilentlyContinue
                
                if ($User) {
                    $AllPrivilegedUsers += [PSCustomObject]@{
                        GroupName         = "Schema Admins"
                        Domain            = $RootDomain
                        MemberName        = $User.Name
                        SamAccountName    = $User.SamAccountName
                        DisplayName       = $User.DisplayName
                        Email             = $User.EmailAddress
                        Title             = $User.Title
                        Department        = $User.Department
                        Enabled           = $User.Enabled
                        MemberCreated     = Format-DateForJSON $User.WhenCreated
                        GroupModified     = Format-DateForJSON $SAGroup.WhenChanged
                        DistinguishedName = $User.DistinguishedName
                    }
                }
            } catch { }
        }
    } catch { }

    # 3. Get AD-Enterprise Systems Admins (Viacom forest only)
    $isViacom = ($ForestName -and ($ForestName.ToLower() -eq "viacom")) -or ($RootDomain -and ($RootDomain.ToLower() -eq "ad.viacom.com"))
    $includeESA = $isViacom -and ((-not $PrivilegedGroupsFromServer) -or ($PrivilegedGroupsFromServer -contains "AD-Enterprise Systems Admins"))
    if ($includeESA) {
      try {
        # Search for the custom group across forest domains (name varies by environment)
        $domainsToSearchForESA = @()
        if ($Forest -and $Forest.Domains) { $domainsToSearchForESA = @($Forest.Domains) } else { $domainsToSearchForESA = @($RootDomain) }
        
        # Try multiple possible group names
        $GroupNames = @(
            "AD-Enterprise Systems Admins",
            "AD-Enterprise-Systems Admins",
            "AD-Enterprise-System Admin",
            "AD-Enterprise System Admins",
            "Enterprise Systems Admins"
        )
        
        $ESAGroup = $null
        $FoundGroupName = $null
        $FoundDomain = $null
        
        foreach ($DomainToSearch in $domainsToSearchForESA) {
            foreach ($GroupName in $GroupNames) {
                try {
                    $ESAGroup = Get-ADGroup -Filter "Name -eq '$GroupName'" -Server $DomainToSearch -Properties WhenChanged -ErrorAction SilentlyContinue
                    if ($ESAGroup) {
                        $FoundGroupName = $ESAGroup.Name
                        $FoundDomain = $DomainToSearch
                        break
                    }
                } catch { }
            }
            if ($ESAGroup) { break }
        }
        
        # If still not found, try wildcard
        if (-not $ESAGroup) {
            try {
                foreach ($DomainToSearch in $domainsToSearchForESA) {
                    $ESAGroup = Get-ADGroup -Filter "Name -like '*Enterprise*System*Admin*'" -Server $DomainToSearch -Properties WhenChanged -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($ESAGroup) {
                        $FoundGroupName = $ESAGroup.Name
                        $FoundDomain = $DomainToSearch
                        break
                    }
                }
                if ($ESAGroup) {
                    $FoundGroupName = $ESAGroup.Name
                }
            } catch { }
        }
        
        if ($ESAGroup) {
            $ESAMembers = Get-SafeGroupMembers -Group $ESAGroup -Server $FoundDomain
            
            foreach ($Member in $ESAMembers) {
                try {
                    $User = Get-ADUser -Identity $Member.distinguishedName -Server $FoundDomain -Properties DisplayName, EmailAddress, Title, Department, Enabled, WhenCreated -ErrorAction SilentlyContinue
                    
                    if ($User) {
                        $AllPrivilegedUsers += [PSCustomObject]@{
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
                            GroupModified     = Format-DateForJSON $ESAGroup.WhenChanged
                            DistinguishedName = $User.DistinguishedName
                        }
                    }
                } catch { }
            }
        }
      } catch { }
    }

    # 4. Get Domain Admins from ALL domains in forest
    foreach ($DomainName in $Forest.Domains) {
        try {
            $DAGroup = Get-ADGroup -Identity "Domain Admins" -Server $DomainName -Properties WhenChanged -ErrorAction SilentlyContinue
            if ($DAGroup) {
                $DAMembers = Get-SafeGroupMembers -Group $DAGroup -Server $DomainName
                
                foreach ($Member in $DAMembers) {
                    try {
                        $User = Get-ADUser -Identity $Member.distinguishedName -Server $DomainName -Properties DisplayName, EmailAddress, Title, Department, Enabled, WhenCreated -ErrorAction SilentlyContinue
                        
                        if ($User) {
                            $AllPrivilegedUsers += [PSCustomObject]@{
                                GroupName         = "Domain Admins"
                                Domain            = $DomainName
                                MemberName        = $User.Name
                                SamAccountName    = $User.SamAccountName
                                DisplayName       = $User.DisplayName
                                Email             = $User.EmailAddress
                                Title             = $User.Title
                                Department        = $User.Department
                                Enabled           = $User.Enabled
                                MemberCreated     = Format-DateForJSON $User.WhenCreated
                                GroupModified     = Format-DateForJSON $DAGroup.WhenChanged
                                DistinguishedName = $User.DistinguishedName
                            }
                        }
                    } catch { }
                }
            }
        } catch { }
    }

    # 5. Get Built-in Administrators from ALL domains in forest
    foreach ($DomainName in $Forest.Domains) {
        try {
            $BAGroup = Get-ADGroup -Identity "Administrators" -Server $DomainName -Properties WhenChanged -ErrorAction SilentlyContinue
            if ($BAGroup) {
                $BAMembers = Get-SafeGroupMembers -Group $BAGroup -Server $DomainName
                
                foreach ($Member in $BAMembers) {
                    try {
                        $User = Get-ADUser -Identity $Member.distinguishedName -Server $DomainName -Properties DisplayName, EmailAddress, Title, Department, Enabled, WhenCreated -ErrorAction SilentlyContinue
                        
                        if ($User) {
                            $AllPrivilegedUsers += [PSCustomObject]@{
                                GroupName         = "Built-in Administrators"
                                Domain            = $DomainName
                                MemberName        = $User.Name
                                SamAccountName    = $User.SamAccountName
                                DisplayName       = $User.DisplayName
                                Email             = $User.EmailAddress
                                Title             = $User.Title
                                Department        = $User.Department
                                Enabled           = $User.Enabled
                                MemberCreated     = Format-DateForJSON $User.WhenCreated
                                GroupModified     = Format-DateForJSON $BAGroup.WhenChanged
                                DistinguishedName = $User.DistinguishedName
                            }
                        }
                    } catch { }
                }
            }
        } catch { }
    }
}
catch {
    @{ Error = "Failed to query forest: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($AllPrivilegedUsers) | ConvertTo-Json -Depth 3
