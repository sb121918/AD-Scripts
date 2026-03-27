param()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $Contacts = Get-ADObject -Filter "objectClass -eq 'contact' -and mail -like '*'" -Properties DisplayName, mail, Description, Company, Department, telephoneNumber, WhenCreated, WhenChanged
    
    $Results = foreach ($Contact in $Contacts) {
        [PSCustomObject]@{
            Name              = $Contact.Name
            DisplayName       = $Contact.DisplayName
            Email             = $Contact.mail
            Description       = $Contact.Description
            Company           = $Contact.Company
            Department        = $Contact.Department
            Phone             = $Contact.telephoneNumber
            Created           = $Contact.WhenCreated
            Modified          = $Contact.WhenChanged
            DistinguishedName = $Contact.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
