param(
    [Parameter(Mandatory=$true)]
    [string]$GroupDN,
    
    [Parameter(Mandatory=$false)]
    [string]$GroupDomain = ""
)

$ErrorActionPreference = 'Stop'

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
    
    # Extract domain from group DN if not provided
    if (-not $GroupDomain -or $GroupDomain -eq "") {
        $GroupDomainParts = ($GroupDN -split ',DC=' | Select-Object -Skip 1)
        $GroupDomain = $GroupDomainParts -join '.'
    }
    
    $Group = $null
    $Members = @()
    
    # Try to get group details
    try {
        $Group = Get-ADGroup -Identity $GroupDN -Server $GroupDomain -Properties Description, mail, ManagedBy, GroupCategory, GroupScope, Created, Modified, Member -ErrorAction Stop
    } catch {
        try {
            $Group = Get-ADGroup -Identity $GroupDN -Properties Description, mail, ManagedBy, GroupCategory, GroupScope, Created, Modified, Member -ErrorAction Stop
        } catch {
            @{
                Error = "Group not found: $GroupDN"
                GroupName = ($GroupDN -split ',')[0].Replace("CN=","")
                TotalMembers = 0
                Members = @()
            } | ConvertTo-Json -Depth 4
            exit 0
        }
    }
    
    # Get group members (without -Recursive to avoid LDAP referral errors)
    if ($Group.Member -and @($Group.Member).Count -gt 0) {
        foreach ($MemberDN in $Group.Member) {
            try {
                # Extract domain from member DN
                $MemberDomainParts = ($MemberDN -split ',DC=' | Select-Object -Skip 1)
                $MemberDomain = $MemberDomainParts -join '.'
                if (-not $MemberDomain) { $MemberDomain = $GroupDomain }
                
                $ADObject = $null
                try {
                    $ADObject = Get-ADObject -Identity $MemberDN -Server $MemberDomain -Properties objectClass, DisplayName, mail, SamAccountName, Enabled, Description -ErrorAction Stop
                } catch {
                    try {
                        $ADObject = Get-ADObject -Identity $MemberDN -Properties objectClass, DisplayName, mail, SamAccountName, Enabled, Description -ErrorAction SilentlyContinue
                    } catch { }
                }
                
                if ($ADObject) {
                    $ObjectType = switch ($ADObject.objectClass) {
                        { $_ -contains 'user' } { 'User' }
                        { $_ -contains 'computer' } { 'Computer' }
                        { $_ -contains 'group' } { 'Group' }
                        { $_ -contains 'contact' } { 'Contact' }
                        default { 'Unknown' }
                    }
                    
                    # Get additional details based on object type
                    $Name = $ADObject.DisplayName
                    $Email = $ADObject.mail
                    $Enabled = $null
                    $Title = $null
                    $Department = $null
                    
                    if ($ObjectType -eq 'User') {
                        try {
                            $UserDetails = Get-ADUser -Identity $MemberDN -Server $MemberDomain -Properties Title, Department, Enabled -ErrorAction SilentlyContinue
                            if ($UserDetails) {
                                $Enabled = $UserDetails.Enabled
                                $Title = $UserDetails.Title
                                $Department = $UserDetails.Department
                            }
                        } catch { }
                    }
                    
                    $Members += [PSCustomObject]@{
                        Name = if ($Name) { $Name } else { ($MemberDN -split ',')[0].Replace("CN=","") }
                        SamAccountName = $ADObject.SamAccountName
                        Type = $ObjectType
                        Email = $Email
                        Enabled = $Enabled
                        Title = $Title
                        Department = $Department
                        Description = $ADObject.Description
                        DistinguishedName = $MemberDN
                        Domain = $MemberDomain
                    }
                } else {
                    # Fallback
                    $Members += [PSCustomObject]@{
                        Name = ($MemberDN -split ',')[0].Replace("CN=","")
                        SamAccountName = $null
                        Type = "Unknown"
                        Email = $null
                        Enabled = $null
                        Title = $null
                        Department = $null
                        Description = $null
                        DistinguishedName = $MemberDN
                        Domain = $MemberDomain
                    }
                }
            } catch {
                $Members += [PSCustomObject]@{
                    Name = ($MemberDN -split ',')[0].Replace("CN=","")
                    SamAccountName = $null
                    Type = "Unknown"
                    Email = $null
                    Enabled = $null
                    Title = $null
                    Department = $null
                    Description = $null
                    DistinguishedName = $MemberDN
                    Domain = "Unknown"
                }
            }
        }
        
        $Members = $Members | Sort-Object Type, Name
    }
    
    # Get ManagedBy name
    $ManagedByName = $null
    if ($Group.ManagedBy) {
        $ManagedByName = ($Group.ManagedBy -split ',')[0].Replace("CN=","")
    }
    
    @{
        GroupName = $Group.Name
        GroupDN = $Group.DistinguishedName
        Description = $Group.Description
        Email = $Group.mail
        ManagedBy = $ManagedByName
        GroupCategory = [string]$Group.GroupCategory
        GroupScope = [string]$Group.GroupScope
        Created = Format-DateForJSON $Group.Created
        Modified = Format-DateForJSON $Group.Modified
        TotalMembers = $Members.Count
        Members = @($Members)
    } | ConvertTo-Json -Depth 4
}
catch {
    @{
        Error = "Failed to get group members: $($_.Exception.Message)"
        GroupName = ($GroupDN -split ',')[0].Replace("CN=","")
        TotalMembers = 0
        Members = @()
    } | ConvertTo-Json -Depth 4
    exit 0
}
