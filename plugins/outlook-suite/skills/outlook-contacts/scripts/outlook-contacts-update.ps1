param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$FullName = "",
    [string]$Email = "",
    [string]$Phone = "",
    [string]$Mobile = "",
    [string]$Company = "",
    [string]$JobTitle = "",
    [string]$Notes = "",
    [string]$Search = ""
)

# Outlook Contact Update Script
# Usage: .\outlook-contacts-update.ps1 -EntryID "00000000..." -Email "newemail@example.com" (preferred)
# Usage: .\outlook-contacts-update.ps1 -Index 1 -Email "newemail@example.com" (fallback)
# Update multiple: .\outlook-contacts-update.ps1 -EntryID "00000000..." -Phone "555-9999" -Company "New Corp"
# With search filter: .\outlook-contacts-update.ps1 -Index 1 -Search "John" -JobTitle "Director"

$Outlook = $null
$Namespace = $null
$Contacts = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")
    $targetContact = $null

    if ($EntryID) {
        $targetContact = $Namespace.GetItemFromID($EntryID)
        if (-not $targetContact) {
            Write-Host "ERROR: Contact not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-contacts-list.ps1 to find contacts and get EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
        $Contacts = $Namespace.GetDefaultFolder(10)  # olFolderContacts = 10
        $items = $Contacts.Items
        $items.Sort("[LastName]")

        # Build filtered list if search is provided
        $filteredContacts = @()
        foreach ($contact in $items) {
            if ($Search) {
                $searchLower = $Search.ToLower()
                $matchFound = $false

                if ($contact.FullName -and $contact.FullName.ToLower().Contains($searchLower)) {
                    $matchFound = $true
                }
                if ($contact.Email1Address -and $contact.Email1Address.ToLower().Contains($searchLower)) {
                    $matchFound = $true
                }
                if ($contact.CompanyName -and $contact.CompanyName.ToLower().Contains($searchLower)) {
                    $matchFound = $true
                }

                if (-not $matchFound) {
                    continue
                }
            }
            $filteredContacts += $contact
        }

        if ($Index -lt 1 -or $Index -gt $filteredContacts.Count) {
            Write-Host "`nContact at index $Index not found." -ForegroundColor Red
            if ($Search) {
                Write-Host "Use outlook-contacts-list.ps1 -Search '$Search' to see available contacts." -ForegroundColor Gray
            } else {
                Write-Host "Use outlook-contacts-list.ps1 to see available contacts." -ForegroundColor Gray
            }
        } else {
            $targetContact = $filteredContacts[$Index - 1]
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target a contact." -ForegroundColor Red
        Write-Host "Usage: .\outlook-contacts-update.ps1 -EntryID ""00000000..."" -Email ""newemail@example.com""" -ForegroundColor Gray
    }

    if ($targetContact) {
        # Track changes
        $changes = @()

        if ($PSBoundParameters.ContainsKey("FullName")) {
            $oldValue = $targetContact.FullName
            $targetContact.FullName = $FullName
            $changes += "Name: '$oldValue' -> '$FullName'"
        }

        if ($PSBoundParameters.ContainsKey("Email")) {
            $oldValue = $targetContact.Email1Address
            $targetContact.Email1Address = $Email
            $changes += "Email: '$oldValue' -> '$Email'"
        }

        if ($PSBoundParameters.ContainsKey("Phone")) {
            $oldValue = $targetContact.BusinessTelephoneNumber
            $targetContact.BusinessTelephoneNumber = $Phone
            $changes += "Work Phone: '$oldValue' -> '$Phone'"
        }

        if ($PSBoundParameters.ContainsKey("Mobile")) {
            $oldValue = $targetContact.MobileTelephoneNumber
            $targetContact.MobileTelephoneNumber = $Mobile
            $changes += "Mobile: '$oldValue' -> '$Mobile'"
        }

        if ($PSBoundParameters.ContainsKey("Company")) {
            $oldValue = $targetContact.CompanyName
            $targetContact.CompanyName = $Company
            $changes += "Company: '$oldValue' -> '$Company'"
        }

        if ($PSBoundParameters.ContainsKey("JobTitle")) {
            $oldValue = $targetContact.JobTitle
            $targetContact.JobTitle = $JobTitle
            $changes += "Job Title: '$oldValue' -> '$JobTitle'"
        }

        if ($PSBoundParameters.ContainsKey("Notes")) {
            $targetContact.Body = $Notes
            if ([string]::IsNullOrEmpty($Notes)) {
                $changes += "Notes cleared"
            } else {
                $changes += "Notes updated"
            }
        }

        if ($changes.Count -eq 0) {
            Write-Host "`nNo changes specified. Use parameters like -FullName, -Email, -Phone, etc." -ForegroundColor Yellow
        } else {
            $targetContact.Save()

            Write-Host "`n=== CONTACT UPDATED ===" -ForegroundColor Green
            Write-Host "Name: $($targetContact.FullName)" -ForegroundColor Yellow

            if ($targetContact.Email1Address) {
                Write-Host "Email: $($targetContact.Email1Address)" -ForegroundColor Gray
            }

            if ($targetContact.CompanyName) {
                if ($targetContact.JobTitle) {
                    Write-Host "Company: $($targetContact.JobTitle) at $($targetContact.CompanyName)" -ForegroundColor Gray
                } else {
                    Write-Host "Company: $($targetContact.CompanyName)" -ForegroundColor Gray
                }
            }

            Write-Host "`nChanges made:" -ForegroundColor Cyan
            foreach ($change in $changes) {
                Write-Host "  - $change" -ForegroundColor Gray
            }
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Contacts) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Contacts) | Out-Null } catch {}
    }
    if ($Namespace) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
