Import-Module ActiveDirectory
Add-Type -AssemblyName Microsoft.VisualBasic

Write-Host "[INIT] Gathering all Distribution Lists from Active Directory..." -ForegroundColor Cyan

# 1. Get all Distribution Groups with their Email property
$Groups = Get-ADGroup -Filter 'GroupCategory -eq "Distribution"' -Properties mail, Description

if ($null -eq $Groups) {
    Write-Warning "No Distribution Lists found in the current domain."
    exit
}

$FullReport = @()

foreach ($Group in $Groups) {
    Write-Host "[PROCESSING] Group: $($Group.Name)" -ForegroundColor Gray
    
    # Get all members (recursive to catch nested groups)
    try {
        $Members = Get-ADGroupMember -Identity $Group.DistinguishedName -Recursive -ErrorAction Stop
        
        if ($Members.Count -eq 0) {
            # Log empty groups so they aren't missed in the audit
            $FullReport += [PSCustomObject]@{
                "Distribution List Name" = $Group.Name
                "Group Email"            = $Group.mail
                "Member Name"            = "-- EMPTY GROUP --"
                "Member Email"           = "N/A"
                "Member Type"            = "N/A"
                "Member SAM"             = "N/A"
                "Group Description"      = $Group.Description
            }
        }

        foreach ($Member in $Members) {
            # Get member email (requires a separate lookup as Get-ADGroupMember doesn't return attributes like 'mail')
            $MemberDetail = Get-ADObject -Identity $Member.distinguishedName -Properties mail
            
            $FullReport += [PSCustomObject]@{
                "Distribution List Name" = $Group.Name
                "Group Email"            = $Group.mail
                "Member Name"            = $Member.name
                "Member Email"           = $MemberDetail.mail
                "Member Type"            = $Member.objectClass
                "Member SAM"             = $Member.SamAccountName
                "Group Description"      = $Group.Description
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve members for $($Group.Name). It may be a protected group or empty."
    }
}

# --- PROFESSIONAL OUTPUT OPTIONS ---

# Option 1: View Results
$ViewChoice = [Microsoft.VisualBasic.Interaction]::MsgBox("Audit Complete. Processed $($Groups.Count) groups.`nTotal membership rows: $($FullReport.Count)`n`nWould you like to open the interactive Grid View?", "YesNo,Question", "Step 1: Review Membership")

if ($ViewChoice -eq "Yes") {
    $FullReport | Out-GridView -Title "Distribution List Membership Audit - $(Get-Date)"
}

# Option 2: Download Report
$DownloadChoice = [Microsoft.VisualBasic.Interaction]::MsgBox("Would you like to save the detailed CSV report to your Desktop?", "YesNo,Question", "Step 2: Export Data")

if ($DownloadChoice -eq "Yes") {
    $ExportPath = "$env:USERPROFILE\Desktop\Detailed_DL_Audit_$(Get-Date -Format 'yyyyMMdd').csv"
    $FullReport | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "[SUCCESS] Report saved to: $ExportPath" -ForegroundColor Green
    
    # Highlight the file in folder
    explorer.exe /select,$ExportPath
}

Write-Host "[FINISH] Membership audit process completed." -ForegroundColor Cyan