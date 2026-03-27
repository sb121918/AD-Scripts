param(
    [string]$Server = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Build common parameters for domain targeting
    $CommonParams = @{}
    if ($Server) {
        $CommonParams['Server'] = $Server
    }
    
    # Fetch all OUs and find empty ones (no limit)
    $OUs = Get-ADOrganizationalUnit -Filter * -Properties Description, WhenCreated, WhenChanged @CommonParams
    
    $EmptyOUs = @()
    foreach ($OU in $OUs) {
        $ChildCount = (Get-ADObject -SearchBase $OU.DistinguishedName -SearchScope OneLevel -Filter * @CommonParams).Count
        if ($ChildCount -eq 0) {
            $EmptyOUs += [PSCustomObject]@{
                Name                      = $OU.Name
                Description               = $OU.Description
                ChildCount                = 0
                Created                   = $OU.WhenCreated
                Modified                  = $OU.WhenChanged
                DistinguishedName         = $OU.DistinguishedName
            }
        }
    }
    
    @($EmptyOUs) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
