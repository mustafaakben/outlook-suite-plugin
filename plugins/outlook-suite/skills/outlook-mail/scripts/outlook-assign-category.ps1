param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [string]$Category,

    [int]$Days = 7
)

# Outlook Category Assigner
# Usage: .\outlook-assign-category.ps1 -EntryID "00000000..." -Category "Project Alpha" (preferred)
# Usage: .\outlook-assign-category.ps1 -Index 3 -Category "Project Alpha" -Days 7 (fallback)
# Assigns a category to the email at the specified index

$Outlook = $null
$Namespace = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")
    $Inbox = $Namespace.GetDefaultFolder(6)

    # Validate category exists
    $categoryExists = $false
    foreach ($cat in $Namespace.Categories) {
        if ($cat.Name -eq $Category) {
            $categoryExists = $true
            break
        }
    }

    if (-not $categoryExists) {
        Write-Host "`nCategory '$Category' does not exist." -ForegroundColor Red
        Write-Host "Create it first with: .\outlook-create-category.ps1 -Name '$Category'" -ForegroundColor Gray
        Write-Host "`nAvailable categories:" -ForegroundColor Yellow
        foreach ($cat in $Namespace.Categories) {
            Write-Host "  - $($cat.Name)" -ForegroundColor Gray
        }
    } else {
        $targetEmail = $null

        if ($EntryID) {
            $targetEmail = $Namespace.GetItemFromID($EntryID)
            if (-not $targetEmail) {
                Write-Host "ERROR: Email not found with provided EntryID." -ForegroundColor Red
                Write-Host "Tip: Use outlook-find.ps1 to get valid EntryIDs." -ForegroundColor Gray
            }
        } elseif ($Index -gt 0) {
            # Get emails with locale-safe date
            $since = (Get-Date).AddDays(-$Days).ToString("g")
            $filter = "[ReceivedTime] >= '$since'"

            $items = $Inbox.Items.Restrict($filter)
            $items.Sort("[ReceivedTime]", $true)

            $count = 0
            foreach ($email in $items) {
                $count++
                if ($count -eq $Index) {
                    $targetEmail = $email
                    break
                }
            }

            if (-not $targetEmail) {
                Write-Host "`nEmail at index $Index not found in last $Days days." -ForegroundColor Red
                Write-Host "Use outlook-read.ps1 -Days $Days to see available emails." -ForegroundColor Gray
                Write-Host "Tip: Try increasing -Days if the email is older." -ForegroundColor Gray
            }
        } else {
            Write-Host "ERROR: Provide -EntryID or -Index to target an email." -ForegroundColor Red
            Write-Host "Usage: .\outlook-assign-category.ps1 -EntryID ""00000000..."" -Category ""Project Alpha""" -ForegroundColor Gray
        }

        if ($targetEmail) {
            # Get sender email with Exchange fallback
            $senderAddr = $targetEmail.SenderEmailAddress
            if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

            # Assign category (handle existing categories)
            $existingCategories = $targetEmail.Categories
            if ($existingCategories) {
                # Check if already assigned
                $catList = $existingCategories -split ",\s*"
                if ($catList -contains $Category) {
                    Write-Host "`nEmail already has category '$Category'." -ForegroundColor Yellow
                    Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Gray
                    Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
                    Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))" -ForegroundColor Gray
                } else {
                    # Append new category
                    $targetEmail.Categories = "$existingCategories, $Category"
                    $targetEmail.Save()

                    Write-Host "`n=== CATEGORY ASSIGNED ===" -ForegroundColor Green
                    Write-Host "Email: $($targetEmail.Subject)" -ForegroundColor Yellow
                    Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
                    Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))" -ForegroundColor Gray
                    Write-Host "Category: $Category" -ForegroundColor Cyan
                    Write-Host "All categories: $($targetEmail.Categories)" -ForegroundColor Gray
                }
            } else {
                $targetEmail.Categories = $Category
                $targetEmail.Save()

                Write-Host "`n=== CATEGORY ASSIGNED ===" -ForegroundColor Green
                Write-Host "Email: $($targetEmail.Subject)" -ForegroundColor Yellow
                Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
                Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))" -ForegroundColor Gray
                Write-Host "Category: $Category" -ForegroundColor Cyan
            }
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Namespace) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
