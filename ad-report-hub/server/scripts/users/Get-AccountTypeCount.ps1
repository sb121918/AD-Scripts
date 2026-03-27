param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $Users = Get-ADUser -Filter 'Enabled -eq $true' -Server $TargetDomain -Properties EmployeeType, extensionAttribute6, Title, Department

    # Process results for summary
    $Results = $Users | Select-Object `
        Name, 
        SamAccountName, 
        @{Name="EmployeeType_Field"; Expression={$_.EmployeeType}},
        @{Name="EA6_Value"; Expression={$_.extensionAttribute6}},
        Title, 
        Department,
        DistinguishedName

    # Create summary counts
    $EmployeeTypeSummary = $Results | Group-Object EmployeeType_Field | 
        Select-Object @{Name="Type"; Expression={if($_.Name){"$($_.Name)"}else{"(Empty)"}}}, Count |
        Sort-Object Count -Descending

    $EA6Summary = $Results | Group-Object EA6_Value | 
        Select-Object @{Name="Type"; Expression={if($_.Name){"$($_.Name)"}else{"(Empty)"}}}, Count |
        Sort-Object Count -Descending

    $Output = [PSCustomObject]@{
        Domain = $TargetDomain
        TotalEnabledUsers = $Results.Count
        EmployeeTypeCounts = @($EmployeeTypeSummary)
        EA6Counts = @($EA6Summary)
        DetailedUsers = @($Results | Select-Object -First 100)
    }

    $Output | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
