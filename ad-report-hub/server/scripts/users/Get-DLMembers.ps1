param(
    [Parameter(Mandatory=$true)]
    [string]$GroupName,
    [int]$Limit = 500
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    # Search for the group by name, display name, or email
    $Filter = "Name -like '*$GroupName*' -or DisplayName -like '*$GroupName*' -or mail -like '*$GroupName*'"
    
    $Groups = Get-ADGroup -Filter $Filter -Properties mail, DisplayName, GroupCategory | 
              Where-Object { $_.GroupCategory -eq "Distribution" }

    if (-not $Groups) {
        @{ Error = "No distribution list found matching '$GroupName'" } | ConvertTo-Json
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

    # Single group found - get members
    $SelectedGroup = $Groups | Select-Object -First 1
    $GroupObj = Get-ADGroup -Identity $SelectedGroup.DistinguishedName -Properties member
    $MemberList = New-Object System.Collections.Generic.List[PSObject]

    foreach ($MemberDN in $GroupObj.member) {
        if ($MemberList.Count -ge $Limit) { break }
        
        try {
            $Object = Get-ADObject -Identity $MemberDN -Properties Name, Title, Department, mail, objectClass
            
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
        GroupName    = $SelectedGroup.DisplayName
        GroupEmail   = $SelectedGroup.mail
        MemberCount  = $MemberList.Count
        Members      = @($MemberList)
    } | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
