param(
    [int]$Limit = 20,
    [string]$Search = ""
)

# Outlook Contacts List Script
# Usage: .\outlook-contacts-list.ps1
# With limit: .\outlook-contacts-list.ps1 -Limit 50
# Filter by name: .\outlook-contacts-list.ps1 -Search "John"

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
    $Contacts = $Namespace.GetDefaultFolder(10)  # olFolderContacts = 10

    $items = $Contacts.Items
    $items.Sort("[LastName]")

    Write-Host "`n=== OUTLOOK CONTACTS ===" -ForegroundColor Cyan

    if ($Search) {
        Write-Host "(Filtered by: '$Search')" -ForegroundColor Gray
    }

    Write-Host ""

    $count = 0
    $displayed = 0

    foreach ($contact in $items) {
        $count++

        # Skip if search filter doesn't match
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

        $displayed++
        if ($displayed -gt $Limit) { break }

        # Display contact
        $name = if ($contact.FullName) { $contact.FullName } else { "(No Name)" }
        Write-Host "$displayed. $name" -ForegroundColor Yellow
        Write-Host "      EntryID: $($contact.EntryID)" -ForegroundColor DarkGray

        if ($contact.Email1Address) {
            Write-Host "      Email: $($contact.Email1Address)" -ForegroundColor Gray
        }

        if ($contact.BusinessTelephoneNumber) {
            Write-Host "      Work: $($contact.BusinessTelephoneNumber)" -ForegroundColor Gray
        } elseif ($contact.MobileTelephoneNumber) {
            Write-Host "      Mobile: $($contact.MobileTelephoneNumber)" -ForegroundColor Gray
        }

        if ($contact.CompanyName) {
            $jobInfo = $contact.CompanyName
            if ($contact.JobTitle) {
                $jobInfo = "$($contact.JobTitle) at $($contact.CompanyName)"
            }
            Write-Host "      Company: $jobInfo" -ForegroundColor Gray
        }

        Write-Host ""
    }

    if ($displayed -eq 0) {
        if ($Search) {
            Write-Host "No contacts found matching '$Search'." -ForegroundColor Gray
        } else {
            Write-Host "No contacts found." -ForegroundColor Gray
        }
    } else {
        Write-Host "--- Showing $displayed of $count contacts ---" -ForegroundColor Cyan
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
