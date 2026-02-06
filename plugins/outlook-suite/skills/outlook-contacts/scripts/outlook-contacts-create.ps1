param(
    [Parameter(Mandatory=$true)]
    [string]$FullName,

    [string]$Email = "",
    [string]$Phone = "",
    [string]$Mobile = "",
    [string]$HomePhone = "",
    [string]$Company = "",
    [string]$JobTitle = "",
    [string]$Notes = ""
)

# Outlook Contact Create Script
# Usage: .\outlook-contacts-create.ps1 -FullName "John Doe"
# Full: .\outlook-contacts-create.ps1 -FullName "John Doe" -Email "john@example.com" -Phone "555-1234" -Company "Acme Inc" -JobTitle "Manager"

$Outlook = $null
$contact = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    # Create contact item (olContactItem = 2)
    $contact = $Outlook.CreateItem(2)

    $contact.FullName = $FullName

    if ($Email) {
        $contact.Email1Address = $Email
    }

    if ($Phone) {
        $contact.BusinessTelephoneNumber = $Phone
    }

    if ($Mobile) {
        $contact.MobileTelephoneNumber = $Mobile
    }

    if ($HomePhone) {
        $contact.HomeTelephoneNumber = $HomePhone
    }

    if ($Company) {
        $contact.CompanyName = $Company
    }

    if ($JobTitle) {
        $contact.JobTitle = $JobTitle
    }

    if ($Notes) {
        $contact.Body = $Notes
    }

    $contact.Save()

    Write-Host "`n=== CONTACT CREATED ===" -ForegroundColor Green
    Write-Host "Name: $FullName" -ForegroundColor Yellow

    if ($Email) {
        Write-Host "Email: $Email" -ForegroundColor Gray
    }

    if ($Phone) {
        Write-Host "Work Phone: $Phone" -ForegroundColor Gray
    }

    if ($Mobile) {
        Write-Host "Mobile: $Mobile" -ForegroundColor Gray
    }

    if ($HomePhone) {
        Write-Host "Home Phone: $HomePhone" -ForegroundColor Gray
    }

    if ($Company) {
        if ($JobTitle) {
            Write-Host "Company: $JobTitle at $Company" -ForegroundColor Gray
        } else {
            Write-Host "Company: $Company" -ForegroundColor Gray
        }
    }

    Write-Host "`nContact saved to Contacts folder." -ForegroundColor Cyan
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($contact) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($contact) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
