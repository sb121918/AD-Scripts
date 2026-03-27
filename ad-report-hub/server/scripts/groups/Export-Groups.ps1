param(
    [string]$Filter = "all",
    [string]$Format = "json",
    [switch]$IncludeMembers,
    [int]$Limit = 1000,
    [string]$TargetDomain = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $ServerArgs = @{}
    if ($TargetDomain -and $TargetDomain.Trim() -ne "") {
        $ServerArgs.Server = $TargetDomain.Trim()
    }

    # Build filter based on parameter
    $ADFilter = switch ($Filter) {
        "security"     { { GroupCategory -eq "Security" } }
        "distribution" { { GroupCategory -eq "Distribution" } }
        "mail-enabled" { { mail -like "*" } }
        default        { "*" }
    }
    
    $Groups = Get-ADGroup @ServerArgs -Filter $ADFilter -Properties DisplayName, mail, Description, GroupCategory, GroupScope, ManagedBy, WhenCreated, WhenChanged, member |
              Select-Object -First $Limit
    
    $Results = foreach ($Group in $Groups) {
        $GroupObj = [PSCustomObject]@{
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
        
        if ($IncludeMembers) {
            $Members = @($Group.member | ForEach-Object { ($_ -split ',')[0].Replace("CN=","") })
            $GroupObj | Add-Member -NotePropertyName "Members" -NotePropertyValue ($Members -join "; ")
        }
        
        $GroupObj
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
