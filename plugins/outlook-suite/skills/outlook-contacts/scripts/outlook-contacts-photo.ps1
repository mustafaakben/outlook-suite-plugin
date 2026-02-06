param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$PhotoPath = "",
    [switch]$Remove,
    [string]$Search = ""
)

# Outlook Contact Photo Script
# Add or remove contact photo
# Usage: .\outlook-contacts-photo.ps1 -EntryID "00000000..." -PhotoPath "C:\Photos\john.jpg" (preferred)
# Usage: .\outlook-contacts-photo.ps1 -Index 1 -PhotoPath "C:\Photos\john.jpg" (fallback)
# Remove photo: .\outlook-contacts-photo.ps1 -EntryID "00000000..." -Remove
# With search: .\outlook-contacts-photo.ps1 -Index 1 -Search "John" -PhotoPath "C:\photo.jpg"

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
        Write-Host "Usage: .\outlook-contacts-photo.ps1 -EntryID ""00000000..."" -PhotoPath ""C:\photo.jpg""" -ForegroundColor Gray
    }

    if ($targetContact) {
        $contactName = $targetContact.FullName

        if ($Remove) {
            # Remove existing photo
            if ($targetContact.HasPicture) {
                $targetContact.RemovePicture()
                $targetContact.Save()
                Write-Host "`n=== PHOTO REMOVED ===" -ForegroundColor Green
                Write-Host "Contact: $contactName" -ForegroundColor Yellow
                Write-Host "Photo has been removed." -ForegroundColor Gray
            } else {
                Write-Host "`nContact '$contactName' does not have a photo." -ForegroundColor Yellow
            }
        } elseif (-not $PhotoPath) {
            Write-Host "`nNo action specified. Use -PhotoPath to add a photo or -Remove to remove." -ForegroundColor Yellow
            Write-Host "Example: .\outlook-contacts-photo.ps1 -EntryID ""..."" -PhotoPath 'C:\Photos\photo.jpg'" -ForegroundColor Gray
        } else {
            # Add photo
            if (-not (Test-Path $PhotoPath)) {
                Write-Host "`nPhoto file not found: $PhotoPath" -ForegroundColor Red
            } else {
                # Validate file type
                $extension = [System.IO.Path]::GetExtension($PhotoPath).ToLower()
                $validExtensions = @(".jpg", ".jpeg", ".png", ".gif", ".bmp")
                if ($extension -notin $validExtensions) {
                    Write-Host "`nInvalid photo format: $extension" -ForegroundColor Red
                    Write-Host "Supported formats: JPG, JPEG, PNG, GIF, BMP" -ForegroundColor Gray
                } else {
                    # Check if contact already has a photo
                    $hadPhoto = $targetContact.HasPicture
                    if ($hadPhoto) {
                        $targetContact.RemovePicture()
                    }

                    # Add the new photo
                    $targetContact.AddPicture($PhotoPath)
                    $targetContact.Save()

                    $fileSize = [math]::Round((Get-Item $PhotoPath).Length / 1KB, 1)
                    $fileName = [System.IO.Path]::GetFileName($PhotoPath)

                    Write-Host "`n=== PHOTO ADDED ===" -ForegroundColor Green
                    Write-Host "Contact: $contactName" -ForegroundColor Yellow
                    Write-Host "Photo: $fileName ($fileSize KB)" -ForegroundColor Gray

                    if ($hadPhoto) {
                        Write-Host "Previous photo was replaced." -ForegroundColor Gray
                    }
                }
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
