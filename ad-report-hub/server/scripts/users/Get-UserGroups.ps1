param(
    [Parameter(Mandatory=$true)]
    [string]$SamAccountName,
    
    [Parameter(Mandatory=$false)]
    [string]$UserDomain = ""
)

$ErrorActionPreference = 'Stop'

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $User = $null
    $FoundDomain = ""
    
    # Try to find user in specified domain first
    if ($UserDomain -and $UserDomain.Trim() -ne "") {
        try {
            $User = Get-ADUser -Identity $SamAccountName -Server $UserDomain -Properties MemberOf, DistinguishedName -ErrorAction Stop
            $FoundDomain = $UserDomain
        } catch { 
            $User = $null
        }
    }
    
    # If not found, search in each domain of the forest
    if (-not $User) {
        try {
            $Forest = Get-ADForest -ErrorAction Stop
            foreach ($Domain in $Forest.Domains) {
                try {
                    $User = Get-ADUser -Identity $SamAccountName -Server $Domain -Properties MemberOf, DistinguishedName -ErrorAction Stop
                    if ($User) {
                        $FoundDomain = $Domain
                        break
                    }
                } catch { 
                    continue 
                }
            }
        } catch {
            # Forest query failed, try current domain
            try {
                $User = Get-ADUser -Identity $SamAccountName -Properties MemberOf, DistinguishedName -ErrorAction Stop
                $FoundDomain = (Get-ADDomain).DNSRoot
            } catch { }
        }
    }
    
    if ($null -eq $User) {
        @{ 
            User = $SamAccountName
            TotalGroups = 0
            Groups = @()
            Error = "User '$SamAccountName' not found"
        } | ConvertTo-Json -Depth 4
        exit 0
    }
    
    # Determine user's domain from DN if not set
    if (-not $FoundDomain -or $FoundDomain -eq "") {
        $FoundDomain = ($User.DistinguishedName -split ',DC=' | Select-Object -Skip 1) -join '.'
    }
    
    $Groups = @()
    
    if ($User.MemberOf -and @($User.MemberOf).Count -gt 0) {
        foreach ($GroupDN in $User.MemberOf) {
            try {
                # Extract domain from group DN
                $GroupDomainParts = ($GroupDN -split ',DC=' | Select-Object -Skip 1)
                $GroupDomain = $GroupDomainParts -join '.'
                
                if (-not $GroupDomain) { $GroupDomain = $FoundDomain }
                
                $Group = $null
                try {
                    $Group = Get-ADGroup -Identity $GroupDN -Server $GroupDomain -Properties GroupCategory, GroupScope, mail, Description, ManagedBy -ErrorAction Stop
                } catch {
                    # Try without specifying server
                    try {
                        $Group = Get-ADGroup -Identity $GroupDN -Properties GroupCategory, GroupScope, mail, Description, ManagedBy -ErrorAction SilentlyContinue
                    } catch { }
                }
                
                if ($Group) {
                    $Groups += [PSCustomObject]@{
                        Name = $Group.Name
                        Type = [string]$Group.GroupCategory
                        Scope = [string]$Group.GroupScope
                        Email = $Group.mail
                        Description = $Group.Description
                        ManagedBy = if ($Group.ManagedBy) { ($Group.ManagedBy -split ',')[0].Replace("CN=","") } else { $null }
                        DistinguishedName = $Group.DistinguishedName
                        Domain = $GroupDomain
                    }
                } else {
                    # Fallback: just extract name from DN
                    $Groups += [PSCustomObject]@{
                        Name = ($GroupDN -split ',')[0].Replace("CN=","")
                        Type = "Unknown"
                        Scope = "Unknown"
                        Email = $null
                        Description = $null
                        ManagedBy = $null
                        DistinguishedName = $GroupDN
                        Domain = $GroupDomain
                    }
                }
            } catch {
                # Last resort fallback
                $Groups += [PSCustomObject]@{
                    Name = ($GroupDN -split ',')[0].Replace("CN=","")
                    Type = "Unknown"
                    Scope = "Unknown"
                    Email = $null
                    Description = $null
                    ManagedBy = $null
                    DistinguishedName = $GroupDN
                    Domain = "Unknown"
                }
            }
        }
        
        $Groups = $Groups | Sort-Object Name
    }
    
    @{
        User = $SamAccountName
        TotalGroups = $Groups.Count
        Groups = @($Groups)
    } | ConvertTo-Json -Depth 4
}
catch {
    @{ 
        User = $SamAccountName
        TotalGroups = 0
        Groups = @()
        Error = "Failed to load groups: $($_.Exception.Message)" 
    } | ConvertTo-Json -Depth 4
    exit 0
}
