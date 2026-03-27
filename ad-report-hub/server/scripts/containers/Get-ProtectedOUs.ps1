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
    
    # Fetch all protected OUs (no limit)
    $OUs = Get-ADOrganizationalUnit -Filter { ProtectedFromAccidentalDeletion -eq $true } -Properties Description, WhenCreated, WhenChanged, ProtectedFromAccidentalDeletion @CommonParams
    
    $Results = foreach ($OU in $OUs) {
        [PSCustomObject]@{
            Name                      = $OU.Name
            Description               = $OU.Description
            Protected                 = $true
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
