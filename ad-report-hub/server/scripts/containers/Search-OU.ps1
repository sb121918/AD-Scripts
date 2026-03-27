param(
    [Parameter(Mandatory=$true)]
    [string]$SearchValue
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $WildcardSearch = "*$($SearchValue.Trim('*'))*"
    
    # Search OUs (no limit)
    $OUs = Get-ADOrganizationalUnit -Filter "Name -like '$WildcardSearch'" -Properties Description, WhenCreated, WhenChanged, ProtectedFromAccidentalDeletion
    
    $Results = foreach ($OU in $OUs) {
        [PSCustomObject]@{
            Name                      = $OU.Name
            Description               = $OU.Description
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
