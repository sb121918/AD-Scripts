# Get-OUChildren.ps1
# Returns users, groups, and computers in a specific OU (by DN)
param(
    [Parameter(Mandatory=$true)]
    [string]$DN,
    [string]$Server = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $results = @{
        Users = @()
        Groups = @()
        Computers = @()
    }

    $getADParams = @{ ErrorAction = 'SilentlyContinue' }
    if ($Server) { $getADParams.Server = $Server }

    # Get users in this OU (one level only)
    $users = Get-ADUser -SearchBase $DN -SearchScope OneLevel -Filter * -Properties DisplayName, mail, Department, Enabled @getADParams
    $results.Users = @($users | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.DisplayName
            SamAccountName = $_.SamAccountName
            Email = $_.mail
            Department = $_.Department
            Enabled = $_.Enabled
        }
    })

    # Get groups in this OU (one level only)
    $groups = Get-ADGroup -SearchBase $DN -SearchScope OneLevel -Filter * -Properties Description, GroupCategory, GroupScope @getADParams
    $results.Groups = @($groups | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.Name
            Description = $_.Description
            Category = [string]$_.GroupCategory
            Scope = [string]$_.GroupScope
        }
    })

    # Get computers in this OU (one level only)
    $computers = Get-ADComputer -SearchBase $DN -SearchScope OneLevel -Filter * -Properties OperatingSystem, Enabled @getADParams
    $results.Computers = @($computers | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.Name
            OperatingSystem = $_.OperatingSystem
            Enabled = $_.Enabled
        }
    })

    $results | ConvertTo-Json -Depth 5
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
}
