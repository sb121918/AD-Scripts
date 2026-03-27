param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $Users = Get-ADUser -Filter 'Enabled -eq $true' -Server $TargetDomain -Properties extensionAttribute6

    $Summary = $Users | Group-Object { if ($_.extensionAttribute6) { $_.extensionAttribute6 } else { '(Empty)' } } |
        Select-Object @{Name="Type"; Expression={$_.Name}}, Count |
        Sort-Object Count -Descending

    $Output = [PSCustomObject]@{
        Domain            = $TargetDomain
        TotalEnabledUsers = $Users.Count
        ReportType        = "EA6 Value"
        Counts            = @($Summary)
    }

    $Output | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
