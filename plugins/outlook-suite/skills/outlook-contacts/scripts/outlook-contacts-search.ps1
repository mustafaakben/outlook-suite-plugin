param(
    [string]$Name = "",
    [string]$Email = "",
    [string]$Company = "",
    [int]$Limit = 10
)

# Outlook Contact Search Script
# Usage: .\outlook-contacts-search.ps1 -Name "John"
# Search by email: .\outlook-contacts-search.ps1 -Email "example.com"
# Search by company: .\outlook-contacts-search.ps1 -Company "Acme"

$Outlook = $null
$Namespace = $null
$Contacts = $null

try {
    if (-not $Name -and -not $Email -and -not $Company) {
        Write-Host "`nPlease specify at least one search parameter:" -ForegroundColor Red
        Write-Host "  -Name ""John""" -ForegroundColor Gray
        Write-Host "  -Email ""example.com""" -ForegroundColor Gray
        Write-Host "  -Company ""Acme""" -ForegroundColor Gray
    } else {
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

        Write-Host "`n=== CONTACT SEARCH RESULTS ===" -ForegroundColor Cyan

        $searchTerms = @()
        if ($Name) { $searchTerms += "Name: '$Name'" }
        if ($Email) { $searchTerms += "Email: '$Email'" }
        if ($Company) { $searchTerms += "Company: '$Company'" }
        Write-Host "Search: $($searchTerms -join ', ')" -ForegroundColor Gray
        Write-Host ""

        $count = 0

        foreach ($contact in $items) {
            $matchName = $true
            $matchEmail = $true
            $matchCompany = $true

            if ($Name) {
                $matchName = $contact.FullName -and $contact.FullName.ToLower().Contains($Name.ToLower())
            }

            if ($Email) {
                $matchEmail = $contact.Email1Address -and $contact.Email1Address.ToLower().Contains($Email.ToLower())
            }

            if ($Company) {
                $matchCompany = $contact.CompanyName -and $contact.CompanyName.ToLower().Contains($Company.ToLower())
            }

            if ($matchName -and $matchEmail -and $matchCompany) {
                $count++
                if ($count -gt $Limit) { break }

                $name = if ($contact.FullName) { $contact.FullName } else { "(No Name)" }
                Write-Host "$count. $name" -ForegroundColor Yellow
                Write-Host "      EntryID: $($contact.EntryID)" -ForegroundColor DarkGray

                if ($contact.Email1Address) {
                    Write-Host "      Email: $($contact.Email1Address)" -ForegroundColor Gray
                }

                if ($contact.BusinessTelephoneNumber) {
                    Write-Host "      Work: $($contact.BusinessTelephoneNumber)" -ForegroundColor Gray
                }

                if ($contact.MobileTelephoneNumber) {
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
        }

        if ($count -eq 0) {
            Write-Host "No contacts found matching your search." -ForegroundColor Gray
        } else {
            Write-Host "--- Found $count contacts ---" -ForegroundColor Cyan
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
