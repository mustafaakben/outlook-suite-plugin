param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$Path = "",
    [string]$Search = ""
)

# Outlook Contact Export Script
# Export contact as vCard (.vcf)
# Usage: .\outlook-contacts-export.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-contacts-export.ps1 -Index 1 (fallback)
# Custom path: .\outlook-contacts-export.ps1 -EntryID "00000000..." -Path "C:\Contacts\john.vcf"
# With search: .\outlook-contacts-export.ps1 -Index 1 -Search "John"

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
        Write-Host "Usage: .\outlook-contacts-export.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if ($targetContact) {
        $contactFullName = [string]$targetContact.FullName
        $contactName = if ([string]::IsNullOrWhiteSpace($contactFullName)) { "(No Name)" } else { $contactFullName }
        $pathLooksLikeDirectory = $Path -and ($Path.Trim().EndsWith("\") -or $Path.Trim().EndsWith("/"))

        # Generate filename if path not provided or is a directory
        if (-not $Path -or (Test-Path $Path -PathType Container) -or $pathLooksLikeDirectory) {
            $folder = if ($Path) { $Path } else { [Environment]::GetFolderPath("UserProfile") + "\Downloads" }

            # Clean name for filename
            $cleanName = $contactFullName -replace '[\\/:*?"<>|]', '_'
            $cleanName = [string]$cleanName
            if (-not [string]::IsNullOrWhiteSpace($cleanName)) {
                $cleanName = $cleanName.Substring(0, [Math]::Min(50, $cleanName.Length))
            } else {
                $cleanName = "contact"
            }

            $fileName = $cleanName + ".vcf"
            $Path = Join-Path $folder $fileName
        }

        # Ensure directory exists
        $directory = [System.IO.Path]::GetDirectoryName($Path)
        if ([string]::IsNullOrWhiteSpace($directory)) {
            $directory = (Get-Location).Path
            $Path = Join-Path $directory ([System.IO.Path]::GetFileName($Path))
        }
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }

        # Ensure .vcf extension before duplicate checks
        if (-not $Path.ToLower().EndsWith(".vcf")) {
            $Path = $Path + ".vcf"
        }

        # Handle duplicate filenames
        $basePath = [System.IO.Path]::GetDirectoryName($Path)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $extension = [System.IO.Path]::GetExtension($Path)
        $counter = 1
        while (Test-Path $Path) {
            $Path = Join-Path $basePath "${baseName}_${counter}${extension}"
            $counter++
        }

        # Export using SaveAs with vCard format (olVCard = 6)
        $targetContact.SaveAs($Path, 6)

        $fileSize = [math]::Round((Get-Item $Path).Length / 1KB, 1)

        Write-Host "`n=== CONTACT EXPORTED ===" -ForegroundColor Green
        Write-Host "Name: $contactName" -ForegroundColor Yellow

        if ($targetContact.Email1Address) {
            Write-Host "Email: $($targetContact.Email1Address)" -ForegroundColor Gray
        }

        if ($targetContact.CompanyName) {
            Write-Host "Company: $($targetContact.CompanyName)" -ForegroundColor Gray
        }

        Write-Host "`nFormat: vCard (.vcf)" -ForegroundColor Cyan
        Write-Host "Size: $fileSize KB"
        Write-Host "Path: $Path" -ForegroundColor Gray
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
