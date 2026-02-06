[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [string]$Name = ""
)

# Outlook Contact Delete Script
# Usage: .\outlook-contacts-delete.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-contacts-delete.ps1 -Index 1 (fallback)
# Delete by name: .\outlook-contacts-delete.ps1 -Name "John Doe"

$Outlook = $null
$Namespace = $null
$Contacts = $null

try {
    if (-not $EntryID -and $Index -eq 0 -and -not $Name) {
        Write-Host "`nPlease specify -EntryID, -Index, or -Name to delete a contact:" -ForegroundColor Red
        Write-Host "  -EntryID ""00000000...""" -ForegroundColor Gray
        Write-Host "  -Index 1" -ForegroundColor Gray
        Write-Host "  -Name ""John Doe""" -ForegroundColor Gray
    } else {
        # Connect to Outlook - try active instance first
        try {
            $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            $Outlook = New-Object -ComObject Outlook.Application
            Start-Sleep -Milliseconds 500
        }

        $Namespace = $Outlook.GetNamespace("MAPI")
        $targetContact = $null
        $contactName = ""

        if ($EntryID) {
            $targetContact = $Namespace.GetItemFromID($EntryID)
            if (-not $targetContact) {
                Write-Host "ERROR: Contact not found with provided EntryID." -ForegroundColor Red
                Write-Host "Tip: Use outlook-contacts-list.ps1 to find contacts and get EntryIDs." -ForegroundColor Gray
            } else {
                $contactName = if ($targetContact.FullName) { $targetContact.FullName } else { "(No Name)" }
            }
        } elseif ($Index -gt 0) {
            $Contacts = $Namespace.GetDefaultFolder(10)  # olFolderContacts = 10
            $items = $Contacts.Items
            $items.Sort("[LastName]")

            $count = 0
            foreach ($contact in $items) {
                $count++
                if ($count -eq $Index) {
                    $targetContact = $contact
                    $contactName = if ($contact.FullName) { $contact.FullName } else { "(No Name)" }
                    break
                }
            }

            if (-not $targetContact) {
                Write-Host "`nContact at index $Index not found." -ForegroundColor Red
                Write-Host "Use outlook-contacts-list.ps1 to see available contacts." -ForegroundColor Gray
            }
        } elseif ($Name) {
            $Contacts = $Namespace.GetDefaultFolder(10)  # olFolderContacts = 10
            $items = $Contacts.Items
            $items.Sort("[LastName]")

            $nameMatches = @()
            foreach ($contact in $items) {
                if ($contact.FullName -and $contact.FullName -eq $Name) {
                    $nameMatches += $contact
                }
            }

            if ($nameMatches.Count -eq 1) {
                $targetContact = $nameMatches[0]
                $contactName = $targetContact.FullName
            } elseif ($nameMatches.Count -gt 1) {
                Write-Host "`nMultiple contacts found with the exact name '$Name'." -ForegroundColor Red
                Write-Host "Use -EntryID or -Index to avoid deleting the wrong contact." -ForegroundColor Gray
                Write-Host "Matching contacts:" -ForegroundColor Cyan
                $matchIndex = 0
                foreach ($match in $nameMatches) {
                    $matchIndex++
                    $matchEmail = if ($match.Email1Address) { $match.Email1Address } else { "(No Email)" }
                    Write-Host "  $matchIndex. $($match.FullName) <$matchEmail> | EntryID: $($match.EntryID)" -ForegroundColor Gray
                }
            } else {
                Write-Host "`nContact '$Name' not found." -ForegroundColor Red
                Write-Host "Use outlook-contacts-search.ps1 -Name ""$Name"" to search." -ForegroundColor Gray
            }
        }

        if ($targetContact) {
            # Store info before deleting
            $email = $targetContact.Email1Address
            $company = $targetContact.CompanyName
            $targetReference = if ($EntryID) {
                "EntryID $EntryID"
            } elseif ($Name) {
                "Name '$Name'"
            } elseif ($Index -gt 0) {
                "Index $Index"
            } else {
                "Selected contact"
            }
            $targetDescription = if ($contactName -and $contactName -ne "(No Name)") {
                "$contactName ($targetReference)"
            } else {
                $targetReference
            }

            if ($PSCmdlet.ShouldProcess($targetDescription, "Delete Outlook contact")) {
                # Delete the contact
                $targetContact.Delete()

                Write-Host "`n=== CONTACT DELETED ===" -ForegroundColor Green
                Write-Host "Name: $contactName" -ForegroundColor Yellow

                if ($email) {
                    Write-Host "Email: $email" -ForegroundColor Gray
                }

                if ($company) {
                    Write-Host "Company: $company" -ForegroundColor Gray
                }
            } else {
                Write-Host "`nDelete cancelled." -ForegroundColor Yellow
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
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
