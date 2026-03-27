param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $Users = Get-ADUser -Filter 'Enabled -eq $true' -Server $TargetDomain -Properties EmployeeType

    $Summary = $Users | Group-Object { if ($_.EmployeeType) { $_.EmployeeType } else { '(Empty)' } } |
        Select-Object @{Name="Type"; Expression={$_.Name}}, Count |
        Sort-Object Count -Descending

    $Output = [PSCustomObject]@{
        Domain            = $TargetDomain
        TotalEnabledUsers = $Users.Count
        ReportType        = "Employee Type"
        Counts            = @($Summary)
    }

    $Output | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
