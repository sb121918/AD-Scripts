param(
    [Parameter(Mandatory=$true)]
    [string]$SearchValue
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $WildcardSearch = "*$($SearchValue.Trim('*'))*"
    
    # Search contacts (no limit)
    $Contacts = Get-ADObject -Filter "objectClass -eq 'contact' -and (Name -like '$WildcardSearch' -or DisplayName -like '$WildcardSearch' -or mail -like '$WildcardSearch')" -Properties DisplayName, mail, Description, Company, Department, telephoneNumber, WhenCreated
    
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
            DistinguishedName = $Contact.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
