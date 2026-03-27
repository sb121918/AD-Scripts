param(
    [Parameter(Mandatory=$false)]
    [string]$GroupName,

    [Parameter(Mandatory=$false)]
    [string]$GroupDN,

    [Parameter(Mandatory=$false)]
    [string]$TargetDomain = "",

    [int]$Limit = 500
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $ServerArgs = @{}
    if ($TargetDomain -and $TargetDomain.Trim() -ne "") {
        $ServerArgs.Server = $TargetDomain.Trim()
    }

    $SelectedGroup = $null

    if ($GroupDN -and $GroupDN.Trim() -ne "") {
        $SelectedGroup = Get-ADGroup @ServerArgs -Identity $GroupDN.Trim() -Properties mail, DisplayName, GroupCategory, member
        if (-not $SelectedGroup) {
            @{ Error = "Group not found for DN provided" } | ConvertTo-Json
            exit 1
        }
        if ($SelectedGroup.GroupCategory -ne "Distribution") {
            @{ Error = "Selected group is not a Distribution Group" } | ConvertTo-Json
            exit 1
        }
    }
    else {
        if (-not $GroupName -or $GroupName.Trim() -eq "") {
            @{ Error = "GroupName or GroupDN is required" } | ConvertTo-Json
            exit 1
        }

        # Search for the group by name, display name, or email
        $SafeGroupName = $GroupName.Trim()
        $Filter = "Name -like '*$SafeGroupName*' -or DisplayName -like '*$SafeGroupName*' -or mail -like '*$SafeGroupName*'"

        $Groups = Get-ADGroup @ServerArgs -Filter $Filter -Properties mail, DisplayName, GroupCategory |
                  Where-Object { $_.GroupCategory -eq "Distribution" }

        if (-not $Groups) {
            @{ Error = "No distribution list found matching '$SafeGroupName'" } | ConvertTo-Json
            exit 1
        }

        # If multiple groups found, return the list for selection
        if (($Groups | Measure-Object).Count -gt 1) {
            $GroupList = $Groups | Select-Object Name, DisplayName, @{Name="Email"; Expression={$_.mail}}, DistinguishedName
            [PSCustomObject]@{
                MultipleResults = $true
                Groups = @($GroupList)
            } | ConvertTo-Json -Depth 5
            exit 0
        }

        $SelectedGroup = $Groups | Select-Object -First 1
        $SelectedGroup = Get-ADGroup @ServerArgs -Identity $SelectedGroup.DistinguishedName -Properties mail, DisplayName, member
    }

    # Get members
    $MemberList = New-Object System.Collections.Generic.List[PSObject]

    foreach ($MemberDN in $SelectedGroup.member) {
        if ($MemberList.Count -ge $Limit) { break }

        try {
            # First try scoping to the requested domain/server
            $Object = $null
            try {
                $Object = Get-ADObject @ServerArgs -Identity $MemberDN -Properties Name, Title, Department, mail, objectClass
            } catch {
                # Fallback: allow ADWS to resolve in its home domain (helps with cross-domain members)
                $Object = Get-ADObject -Identity $MemberDN -Properties Name, Title, Department, mail, objectClass -ErrorAction Stop
            }

            $MemberList.Add([PSCustomObject]@{
                Name       = $Object.Name
                Type       = $Object.objectClass
                Title      = if ($Object.Title) { $Object.Title } else { "N/A" }
                Department = if ($Object.Department) { $Object.Department } else { "N/A" }
                Email      = if ($Object.mail) { $Object.mail } else { "N/A" }
                DN         = $Object.DistinguishedName
            })
        } catch {
            $MemberList.Add([PSCustomObject]@{
                Name       = "Unresolved/Foreign Member"
                Type       = "Unknown"
                Title      = "N/A"
                Department = "N/A"
                Email      = $MemberDN
                DN         = "Error Resolving"
            })
        }
    }

    [PSCustomObject]@{
        GroupName   = $SelectedGroup.DisplayName
        GroupEmail  = $SelectedGroup.mail
        GroupDN     = $SelectedGroup.DistinguishedName
        MemberCount = $MemberList.Count
        Members     = @($MemberList)
    } | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
