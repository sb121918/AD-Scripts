param(
    [Parameter(Mandatory=$true)]
    [string]$GroupName,

    [Parameter(Mandatory=$false)]
    [string]$TargetDomain = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $ServerArgs = @{}
    if ($TargetDomain -and $TargetDomain.Trim() -ne "") {
        $ServerArgs.Server = $TargetDomain.Trim()
    }

    $Group = Get-ADGroup @ServerArgs -Identity $GroupName -Properties *
    
    if ($null -eq $Group) {
        @{ Error = "Group not found" } | ConvertTo-Json
        exit
    }
    
    $Result = [PSCustomObject]@{
        Name              = $Group.Name
        DisplayName       = $Group.DisplayName
        SamAccountName    = $Group.SamAccountName
        Email             = $Group.mail
        Description       = $Group.Description
        Type              = [string]$Group.GroupCategory
        Scope             = [string]$Group.GroupScope
        MemberCount       = @($Group.member).Count
        ManagedBy         = if ($Group.ManagedBy) { ($Group.ManagedBy -split ',')[0].Replace("CN=","") } else { $null }
        ManagedByDN       = $Group.ManagedBy
        Notes             = $Group.info
        Created           = $Group.WhenCreated
        Modified          = $Group.WhenChanged
        DistinguishedName = $Group.DistinguishedName
    }
    
    $Result | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
