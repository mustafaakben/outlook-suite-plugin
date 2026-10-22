param(
    [Parameter(Mandatory = $true)]
    [string]$FullName,

    [string]$Email = "",
    [string]$Email2 = "",
    [string]$Phone = "",
    [string]$Mobile = "",
    [string]$HomePhone = "",
    [string]$Company = "",
    [string]$JobTitle = "",
    [string]$Department = "",
    [string]$Office = "",
    [string]$Categories = "",
    [string]$Notes = "",
    [switch]$Json
)

# Outlook Contact Create: a new contact in the default Contacts folder.
# Usage: .\outlook-contacts-create.ps1 -FullName "John Doe"
# Usage: .\outlook-contacts-create.ps1 -FullName "John Doe" -Email "john@example.com" -Phone "555-1234" -Company "Acme Inc" -JobTitle "Manager" -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

try {
    $null = Connect-Outlook
    $contact = $Outlook.CreateItem(2)   # olContactItem
    $contact.FullName = $FullName
    if ($Email) { $contact.Email1Address = $Email }
    if ($Email2) { $contact.Email2Address = $Email2 }
    if ($Phone) { $contact.BusinessTelephoneNumber = $Phone }
    if ($Mobile) { $contact.MobileTelephoneNumber = $Mobile }
    if ($HomePhone) { $contact.HomeTelephoneNumber = $HomePhone }
    if ($Company) { $contact.CompanyName = $Company }
    if ($JobTitle) { $contact.JobTitle = $JobTitle }
    if ($Department) { $contact.Department = $Department }
    if ($Office) { $contact.OfficeLocation = $Office }
    if ($Categories) { $contact.Categories = $Categories }
    if ($Notes) { $contact.Body = $Notes }
    $contact.Save()
    Write-ContactResult -Title "CONTACT CREATED" -Rec (ConvertTo-ContactRecord $contact) -Action "create" -Json:$Json
}
catch {
    Stop-WithError $_
}
