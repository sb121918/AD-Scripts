param(
    [string]$TargetDomain
)

function Get-NestedGroupMembers {
    param(
        [string]$GroupDN,
        [string]$RootGroup,
        [string]$Path,
        [hashtable]$Visited,
        [string]$Server
    )
    
    $results = @()
    
    if ($Visited.ContainsKey($GroupDN)) { return $results }
    $Visited[$GroupDN] = $true
    
    $serverParam = @{}
    if ($Server) { $serverParam.Server = $Server }
    
    try {
        $members = Get-ADGroupMember -Identity $GroupDN @serverParam -ErrorAction SilentlyContinue
        
        foreach ($member in $members) {
            $currentPath = if ($Path) { "$Path -> $($member.Name)" } else { $member.Name }
            
            if ($member.objectClass -eq 'group') {
                $results += Get-NestedGroupMembers -GroupDN $member.distinguishedName -RootGroup $RootGroup -Path $currentPath -Visited $Visited -Server $Server
            }
            elseif ($member.objectClass -eq 'user') {
                $user = Get-ADUser -Identity $member.distinguishedName @serverParam -Properties DisplayName, mail, Enabled, LastLogonDate, Title, Department
                $results += [PSCustomObject]@{
                    PrivilegedGroup   = $RootGroup
                    MemberName        = $user.DisplayName
                    SamAccountName    = $user.SamAccountName
                    Email             = $user.mail
                    Title             = $user.Title
                    Department        = $user.Department
                    Enabled           = $user.Enabled
                    LastLogon         = $user.LastLogonDate
                    MembershipPath    = $currentPath
                    IsNested          = $Path -ne $null -and $Path -ne ''
                    NestingDepth      = ($currentPath -split ' -> ').Count - 1
                    Domain            = ($member.distinguishedName -split ',DC=' | Select-Object -Skip 1) -join '.'
                    DistinguishedName = $member.distinguishedName
                }
            }
        }
    }
    catch {
        # Skip groups we can't access
    }
    
    return $results
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $WarningPreference = 'SilentlyContinue'

    $anchorDomain = if ($TargetDomain) { $TargetDomain } else { (Get-ADDomain).DNSRoot }
    $forest = Get-ADForest -Server $anchorDomain -ErrorAction Stop
    $domainsToQuery = @($forest.Domains)

    $privilegedGroups = @(
        "Domain Admins",
        "Enterprise Admins",
        "Schema Admins",
        "Administrators",
        "Account Operators",
        "Backup Operators",
        "Server Operators"
    )

    $allResults = @()

    foreach ($domainName in $domainsToQuery) {
        foreach ($groupName in $privilegedGroups) {
            try {
                $group = Get-ADGroup -Filter "Name -eq '$groupName'" -Server $domainName -ErrorAction SilentlyContinue
                if ($group) {
                    $visited = @{}
                    $members = Get-NestedGroupMembers -GroupDN $group.DistinguishedName -RootGroup $groupName -Path "" -Visited $visited -Server $domainName
                    $allResults += $members
                }
            }
            catch { }
        }
    }

    $uniqueResults = $allResults | Sort-Object PrivilegedGroup, SamAccountName, MembershipPath -Unique
    
    @($uniqueResults) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
