param()

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

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$Results = @()

try {
    $Forest = Get-ADForest -ErrorAction Stop
    $RootDomain = $Forest.RootDomain
    
    # Get Enterprise Admins group from root domain
    $Group = Get-ADGroup -Identity "Enterprise Admins" -Server $RootDomain -Properties WhenCreated, WhenChanged, Description -ErrorAction Stop
    
    # Get members without -Recursive to avoid referral issues
    $Members = @()
    try {
        $Members = Get-ADGroupMember -Identity $Group.DistinguishedName -Server $RootDomain -ErrorAction Stop | 
                   Where-Object { $_.objectClass -eq 'user' }
    } catch {
        # Fallback: query users by memberOf
        try {
            $GroupDN = $Group.DistinguishedName
            $Members = Get-ADUser -Filter "memberOf -eq '$GroupDN'" -Server $RootDomain -ErrorAction SilentlyContinue
        } catch { }
    }
    
    foreach ($Member in $Members) {
        try {
            $User = Get-ADUser -Identity $Member.distinguishedName -Server $RootDomain -Properties DisplayName, EmailAddress, Title, Department, Enabled, WhenCreated -ErrorAction SilentlyContinue
            
            if ($User) {
                $Results += [PSCustomObject]@{
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
                    GroupModified     = Format-DateForJSON $Group.WhenChanged
                    DistinguishedName = $User.DistinguishedName
                }
            }
        } catch { }
    }
}
catch {
    @{ Error = "Failed to query Enterprise Admins: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($Results) | ConvertTo-Json -Depth 3
