param(
    [Parameter(Mandatory=$false)]
    [string]$TargetDomain = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $ServerArgs = @{}
    if ($TargetDomain -and $TargetDomain.Trim() -ne "") {
        $ServerArgs.Server = $TargetDomain.Trim()
    }

    # Fetch all security groups (no limit)
    $Groups = Get-ADGroup @ServerArgs -Filter { GroupCategory -eq "Security" } -Properties DisplayName, mail, Description, GroupScope, ManagedBy, WhenCreated, member
    
    $Results = foreach ($Group in $Groups) {
        [PSCustomObject]@{
            Name              = $Group.Name
            DisplayName       = $Group.DisplayName
            Email             = $Group.mail
            Description       = $Group.Description
            Type              = "Security"
            Scope             = [string]$Group.GroupScope
            MemberCount       = @($Group.member).Count
            ManagedBy         = if ($Group.ManagedBy) { ($Group.ManagedBy -split ',')[0].Replace("CN=","") } else { $null }
            Created           = $Group.WhenCreated
            DistinguishedName = $Group.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
