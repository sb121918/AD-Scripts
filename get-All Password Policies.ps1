# --- Configuration ---
$FullReport = New-Object System.Collections.Generic.List[PSObject]
Write-Host "Identifying all domains in the forest..." -ForegroundColor Cyan

try {
    $AllDomains = (Get-ADForest).Domains
} catch {
    Write-Host "Error: Could not retrieve forest. Ensure RSAT is installed." -ForegroundColor Red
    return
}

foreach ($Domain in $AllDomains) {
    Write-Host "Processing Domain: $Domain" -ForegroundColor White
    
    # 1. Get Default Domain Password Policy
    try {
        $DefaultPolicy = Get-ADDefaultDomainPasswordPolicy -Server $Domain
        if ($DefaultPolicy) {
            $FullReport.Add([PSCustomObject]@{
                Domain             = $Domain
                PolicyName         = "DEFAULT DOMAIN POLICY"
                Type               = "Default"
                Precedence         = "N/A"
                Complexity         = $DefaultPolicy.ComplexityEnabled
                MinPasswordLength  = $DefaultPolicy.MinPasswordLength
                PasswordHistory    = $DefaultPolicy.PasswordHistoryCount
                LockoutThreshold   = $DefaultPolicy.LockoutThreshold
                MaxPasswordAgeDays = $DefaultPolicy.MaxPasswordAge.Days
                AppliesTo          = "All Users (Domain)"
            })
        }
    } catch {
        Write-Host "  [!] No Default Policy returned for $Domain" -ForegroundColor Yellow
    }

    # 2. Get Fine-Grained Password Policies (PSOs)
    try {
        # Searching specifically in the Password Settings Container
        $FGPPs = Get-ADFineGrainedPasswordPolicy -Filter * -Server $Domain -Properties *
        
        if ($FGPPs) {
            Write-Host "  [+] Found $($FGPPs.Count) Fine-Grained Policies in $Domain" -ForegroundColor Green
            foreach ($Policy in $FGPPs) {
                $FullReport.Add([PSCustomObject]@{
                    Domain             = $Domain
                    PolicyName         = $Policy.Name
                    Type               = "Fine-Grained (PSO)"
                    Precedence         = $Policy.Precedence
                    Complexity         = $Policy.ComplexityEnabled
                    MinPasswordLength  = $Policy.MinPasswordLength
                    PasswordHistory    = $Policy.PasswordHistoryCount
                    LockoutThreshold   = $Policy.LockoutThreshold
                    MaxPasswordAgeDays = $Policy.MaxPasswordAge.Days
                    AppliesTo          = ($Policy.AppliesTo -join '; ') # Shows the DNs of groups/users
                })
            }
        } else {
            Write-Host "  [-] No Fine-Grained Policies exist in $Domain" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  [!] Error searching Fine-Grained Policies in $Domain" -ForegroundColor Red
    }
}

# --- Output ---
if ($FullReport.Count -gt 0) {
    $FullReport | Out-GridView -Title "Forest Password Policy Audit (Default & Fine-Grained)"
} else {
    Write-Host "No policies found at all." -ForegroundColor Red
}