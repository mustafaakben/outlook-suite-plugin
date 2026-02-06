param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$Category = "",
    [int]$Days = 7,
    [switch]$All
)

# Outlook Remove Category Script
# Usage: .\outlook-remove-category.ps1 -EntryID "00000000..." -Category "Important" (preferred)
# Usage: .\outlook-remove-category.ps1 -Index 1 -Category "Important" (fallback)
# Remove all categories: .\outlook-remove-category.ps1 -EntryID "00000000..." -All

$Outlook = $null
$Namespace = $null

try {
    # Validate parameters early
    if (-not $All -and -not $Category) {
        Write-Host "`nError: Please specify -Category or use -All flag." -ForegroundColor Red
        Write-Host "Usage: .\outlook-remove-category.ps1 -Index 1 -Category `"Important`"" -ForegroundColor Gray
        Write-Host "   or: .\outlook-remove-category.ps1 -Index 1 -All" -ForegroundColor Gray
    } else {
        # Connect to Outlook - try active instance first
        try {
            $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            $Outlook = New-Object -ComObject Outlook.Application
            Start-Sleep -Milliseconds 500
        }

        $Namespace = $Outlook.GetNamespace("MAPI")
        $targetEmail = $null

        if ($EntryID) {
            $targetEmail = $Namespace.GetItemFromID($EntryID)
            if (-not $targetEmail) {
                Write-Host "ERROR: Email not found with provided EntryID." -ForegroundColor Red
                Write-Host "Tip: Use outlook-find.ps1 to get valid EntryIDs." -ForegroundColor Gray
            }
        } elseif ($Index -gt 0) {
            $Inbox = $Namespace.GetDefaultFolder(6)

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
            Write-Host "Usage: .\outlook-remove-category.ps1 -EntryID ""00000000..."" -Category ""Important""" -ForegroundColor Gray
        }

        if ($targetEmail) {
            # Get sender email with Exchange fallback
            $senderAddr = $targetEmail.SenderEmailAddress
            if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

            # Get current categories
            $currentCategories = $targetEmail.Categories

            if (-not $currentCategories) {
                Write-Host "`nEmail has no categories assigned." -ForegroundColor Yellow
                Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Gray
                Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
                Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))" -ForegroundColor Gray
            } elseif ($All) {
                # Remove all categories
                $previousCategories = $currentCategories
                $targetEmail.Categories = ""
                $targetEmail.Save()

                Write-Host "`n=== ALL CATEGORIES REMOVED ===" -ForegroundColor Green
                Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
                Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
                Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))" -ForegroundColor Gray
                Write-Host "Removed: $previousCategories" -ForegroundColor Gray
            } else {
                # Remove specific category
                $catList = $currentCategories -split ",\s*"

                if ($catList -notcontains $Category) {
                    Write-Host "`nCategory '$Category' not found on this email." -ForegroundColor Red
                    Write-Host "Current categories: $currentCategories" -ForegroundColor Gray
                } else {
                    $newCatList = $catList | Where-Object { $_ -ne $Category }
                    $targetEmail.Categories = $newCatList -join ", "
                    $targetEmail.Save()

                    Write-Host "`n=== CATEGORY REMOVED ===" -ForegroundColor Green
                    Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
                    Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
                    Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))" -ForegroundColor Gray
                    Write-Host "Removed: $Category" -ForegroundColor Cyan
                    if ($targetEmail.Categories) {
                        Write-Host "Remaining: $($targetEmail.Categories)" -ForegroundColor Gray
                    } else {
                        Write-Host "No categories remaining." -ForegroundColor Gray
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
    if ($Namespace) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
