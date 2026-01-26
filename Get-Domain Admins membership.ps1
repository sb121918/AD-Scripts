Import-Module ActiveDirectory
Add-Type -AssemblyName Microsoft.VisualBasic

# Initialize Report Metadata
$ScanDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$Results = @()
$Forest = Get-ADForest
$Domains = $Forest.Domains

Write-Host "[INIT] Starting Security Audit for Forest: $($Forest.RootDomain)" -ForegroundColor Cyan

foreach ($Domain in $Domains) {
    # Target PDC Emulator for the most accurate security data
    $DomainInfo = Get-ADDomain -Identity $Domain
    $PDC = $DomainInfo.PDCEmulator
    
    Write-Host "[INFO] Analyzing Privileged Members in: $Domain (Server: $PDC)" -ForegroundColor Gray

    try {
        # Recursive search to catch nested group memberships
        $Members = Get-ADGroupMember -Identity "Domain Admins" -Recursive -Server $PDC
        
        foreach ($Member in $Members) {
            # Retrieve critical security attributes
            $UserDetails = Get-ADUser -Identity $Member.distinguishedName -Server $PDC -Properties PasswordLastSet, PasswordNeverExpires, LastLogonDate, Enabled
            
            # Determine Account Status
            $Status = if($UserDetails.Enabled) { "Active" } else { "Disabled" }
            
            # Construct Professional Report Object
            $Results += [PSCustomObject]@{
                "Organizational Domain" = $Domain.ToUpper()
                "Full Name"             = $UserDetails.Name
                "User Identifier"       = $UserDetails.SamAccountName
                "Account Status"        = $Status
                "Password Last Set"     = $UserDetails.PasswordLastSet
                "Non-Expiring Password" = $UserDetails.PasswordNeverExpires
                "Last Authenticated"    = $UserDetails.LastLogonDate
                "Distinguished Name"    = $UserDetails.DistinguishedName
            }
        }
    }
    catch {
        Write-Warning "[ERROR] Connectivity issue or Access Denied for $Domain."
    }
}

# --- OUTPUT OPTION 1: GRID VIEW ---
$ViewChoice = [Microsoft.VisualBasic.Interaction]::MsgBox("Audit Complete. Found $($Results.Count) members.`n`nWould you like to open the interactive Grid View?", "YesNo,Question", "Step 1: Data Review")

if ($ViewChoice -eq "Yes") {
    $Results | Out-GridView -Title "IAM Review: Privileged Domain Access ($ScanDate)"
}

# --- OUTPUT OPTION 2: DOWNLOAD CSV ---
$DownloadChoice = [Microsoft.VisualBasic.Interaction]::MsgBox("Would you like to download the security report to your Desktop?", "YesNo,Question", "Step 2: Export Report")

if ($DownloadChoice -eq "Yes") {
    $ExportPath = "$env:USERPROFILE\Desktop\Privileged_Access_Report_$(Get-Date -Format 'yyyyMMdd').csv"
    $Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "[SUCCESS] Security Report Generated: $ExportPath" -ForegroundColor Green
    
    # Optional: Open the folder to show the file
    explorer.exe /select,$ExportPath
}

Write-Host "[FINISH] Audit process completed." -ForegroundColor Cyan