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
    
    # Fetch all OUs (no limit)
    $OUs = Get-ADOrganizationalUnit -Filter * -Properties Description, WhenCreated, WhenChanged, ProtectedFromAccidentalDeletion @CommonParams
    
    $Results = foreach ($OU in $OUs) {
        # Count child objects
        $ChildCount = (Get-ADObject -SearchBase $OU.DistinguishedName -SearchScope OneLevel -Filter * @CommonParams).Count
        
        [PSCustomObject]@{
            Name                      = $OU.Name
            Description               = $OU.Description
            ChildCount                = $ChildCount
            Protected                 = $OU.ProtectedFromAccidentalDeletion
            Created                   = $OU.WhenCreated
            Modified                  = $OU.WhenChanged
            DistinguishedName         = $OU.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
