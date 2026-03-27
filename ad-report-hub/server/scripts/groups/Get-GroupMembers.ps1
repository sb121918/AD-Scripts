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

    $Group = Get-ADGroup @ServerArgs -Identity $GroupName -Properties member, mail, Description, ManagedBy
    
    if ($null -eq $Group) {
        @{ Error = "Group not found" } | ConvertTo-Json
        exit
    }
    
    $Members = @()
    foreach ($MemberDN in $Group.member) {
        try {
            $Object = Get-ADObject @ServerArgs -Identity $MemberDN -Properties Name, Title, Department, mail, objectClass, SamAccountName
            
            $Members += [PSCustomObject]@{
                Name              = $Object.Name
                SamAccountName    = $Object.SamAccountName
                Type              = $Object.objectClass
                Title             = $Object.Title
                Department        = $Object.Department
                Email             = $Object.mail
                DistinguishedName = $Object.DistinguishedName
            }
        } catch {
            $Members += [PSCustomObject]@{
                Name              = ($MemberDN -split ',')[0].Replace("CN=","")
                SamAccountName    = $null
                Type              = "Unknown"
                Title             = $null
                Department        = $null
                Email             = $null
                DistinguishedName = $MemberDN
            }
        }
    }
    
    @{
        GroupName = $Group.Name
        Email = $Group.mail
        Description = $Group.Description
        ManagedBy = if ($Group.ManagedBy) { ($Group.ManagedBy -split ',')[0].Replace("CN=","") } else { $null }
        MemberCount = $Members.Count
        Members = @($Members | Sort-Object Name)
    } | ConvertTo-Json -Depth 4
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
