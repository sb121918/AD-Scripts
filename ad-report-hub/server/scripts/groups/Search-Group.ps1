param(
    [Parameter(Mandatory=$true)]
    [string]$SearchValue,

    [Parameter(Mandatory=$false)]
    [string]$TargetDomain = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $ServerArgs = @{}
    if ($TargetDomain -and $TargetDomain.Trim() -ne "") {
        $ServerArgs.Server = $TargetDomain.Trim()
    }

    $WildcardSearch = "*$($SearchValue.Trim('*'))*"
    
    $Filter = "Name -like '$WildcardSearch' -or DisplayName -like '$WildcardSearch' -or mail -like '$WildcardSearch'"
    
    # Search groups (no limit)
    $Groups = Get-ADGroup @ServerArgs -Filter $Filter -Properties DisplayName, mail, Description, GroupCategory, GroupScope, ManagedBy, WhenCreated, WhenChanged, member
    
    if ($null -eq $Groups -or @($Groups).Count -eq 0) {
        @() | ConvertTo-Json
        exit
    }
    
    $Results = foreach ($Group in $Groups) {
        [PSCustomObject]@{
            Name              = $Group.Name
            DisplayName       = $Group.DisplayName
            Email             = $Group.mail
            Description       = $Group.Description
            Type              = [string]$Group.GroupCategory
            Scope             = [string]$Group.GroupScope
            MemberCount       = @($Group.member).Count
            ManagedBy         = if ($Group.ManagedBy) { ($Group.ManagedBy -split ',')[0].Replace("CN=","") } else { $null }
            Created           = $Group.WhenCreated
            Modified          = $Group.WhenChanged
            DistinguishedName = $Group.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
