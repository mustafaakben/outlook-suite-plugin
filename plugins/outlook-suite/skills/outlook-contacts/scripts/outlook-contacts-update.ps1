param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$FullName,
    [string]$Email,
    [string]$Email2,
    [string]$Phone,
    [string]$Mobile,
    [string]$HomePhone,
    [string]$Company,
    [string]$JobTitle,
    [string]$Department,
    [string]$Office,
    [string]$Categories,
    [string]$Notes,
    [string]$Search = "",
    [switch]$Json
)

# Outlook Contact Update: change fields of one contact. Only the fields you pass change;
# pass "" to clear one.
# Usage: .\outlook-contacts-update.ps1 -EntryID "00000000..." -Email "newemail@example.com"
# Usage: .\outlook-contacts-update.ps1 -Index 1 -Search "John" -JobTitle "Director" -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

# parameter -> ContactItem property
$fields = [ordered]@{
    FullName = "FullName"; Email = "Email1Address"; Email2 = "Email2Address"; Phone = "BusinessTelephoneNumber"
    Mobile = "MobileTelephoneNumber"; HomePhone = "HomeTelephoneNumber"; Company = "CompanyName"; JobTitle = "JobTitle"
    Department = "Department"; Office = "OfficeLocation"; Categories = "Categories"; Notes = "Body"
}

$given = @($fields.Keys | Where-Object { $PSBoundParameters.ContainsKey($_) })
if ($given.Count -eq 0) {
    Stop-WithError "No changes given." -Hint "Pass at least one field, e.g. -Email, -Phone, -Mobile, -Company, -JobTitle, -Department, -Notes."
}

try {
    $null = Connect-Outlook
    $contact = Get-ContactTarget -Namespace $Namespace -EntryID $EntryID -Index $Index -Search $Search
    $changes = [ordered]@{}
    foreach ($p in $given) {
        $prop = $fields[$p]
        $new = [string]$PSBoundParameters[$p]
        if ($p -eq "Notes") {
            $contact.Body = $new
            $changes["Notes"] = $(if ($new) { "updated" } else { "cleared" })
            continue
        }
        $old = [string]$contact.$prop
        $contact.$prop = $new
        $changes[$p] = "'$old' -> '$new'"
    }
    $contact.Save()
    Write-ContactResult -Title "CONTACT UPDATED" -Rec (ConvertTo-ContactRecord $contact) -Changes $changes -Action "update" -Json:$Json
}
catch {
    Stop-WithError $_
}
