param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [int]$Days = 7,
    [switch]$Permanent,
    [switch]$Confirm
)

# Outlook Delete Email Script
# Usage: .\outlook-delete.ps1 -EntryID "00000000..." -Confirm (preferred)
# Usage: .\outlook-delete.ps1 -Index 1 -Confirm (fallback)

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
    $targetEmail = $null

    if ($EntryID) {
        $targetEmail = $Namespace.GetItemFromID($EntryID)
        if (-not $targetEmail) {
            Write-Host "ERROR: Email not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-find.ps1 to get valid EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
        $Inbox = $Namespace.GetDefaultFolder(6)
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
        Write-Host "Usage: .\outlook-delete.ps1 -EntryID ""00000000..."" -Confirm" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Get sender email with Exchange fallback
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

        $subject = $targetEmail.Subject
        $senderName = $targetEmail.SenderName
        $date = $targetEmail.ReceivedTime.ToString("g")

        # Show preview
        $deleteType = if ($Permanent) { "PERMANENT DELETE" } else { "DELETE (move to Deleted Items)" }
        Write-Host "`n=== $deleteType PREVIEW ===" -ForegroundColor Yellow
        Write-Host "Subject: $subject" -ForegroundColor White
        Write-Host "From: $senderName <$senderAddr>"
        Write-Host "Date: $date"

        if ($Permanent) {
            Write-Host "`n[!] WARNING: Permanent delete cannot be undone!" -ForegroundColor Red
        }

        if (-not $Confirm) {
            Write-Host "`n[!] Add -Confirm to actually delete this email." -ForegroundColor Cyan
        } else {
            if ($Permanent) {
                $deletedItems = $Namespace.GetDefaultFolder(3)
                $deletedStoreID = $null
                $deletedFolderEntryID = $null
                try { $deletedStoreID = $deletedItems.StoreID } catch {}
                try { $deletedFolderEntryID = $deletedItems.EntryID } catch {}

                $targetStoreID = $null
                $targetFolderEntryID = $null
                try { $targetStoreID = $targetEmail.Parent.StoreID } catch {}
                try { $targetFolderEntryID = $targetEmail.Parent.EntryID } catch {}

                $emailForPermanentDelete = $null
                $idForPermanentDelete = $null
                $storeForPermanentDelete = $null

                $alreadyInDeletedItems = $false
                if ($targetStoreID -and $deletedStoreID -and $targetFolderEntryID -and $deletedFolderEntryID) {
                    $alreadyInDeletedItems = ($targetStoreID -eq $deletedStoreID -and $targetFolderEntryID -eq $deletedFolderEntryID)
                }

                if ($alreadyInDeletedItems) {
                    try { $idForPermanentDelete = $targetEmail.EntryID } catch {}
                    $storeForPermanentDelete = $deletedStoreID
                    $emailForPermanentDelete = $targetEmail
                } else {
                    try {
                        $movedEmail = $targetEmail.Move($deletedItems)
                        if ($movedEmail) {
                            try { $idForPermanentDelete = $movedEmail.EntryID } catch {}
                            try { $storeForPermanentDelete = $movedEmail.Parent.StoreID } catch {}
                            if (-not $storeForPermanentDelete) {
                                $storeForPermanentDelete = $deletedStoreID
                            }
                            $emailForPermanentDelete = $movedEmail
                        }
                    } catch {
                        Write-Host "`n[SAFE-FAIL] Could not move the message to Deleted Items for permanent deletion." -ForegroundColor Yellow
                        Write-Host "No permanent delete was performed." -ForegroundColor Gray
                    }
                }

                if ($idForPermanentDelete -and $storeForPermanentDelete) {
                    try {
                        $resolvedEmail = $Namespace.GetItemFromID($idForPermanentDelete, $storeForPermanentDelete)
                        if ($resolvedEmail) {
                            $emailForPermanentDelete = $resolvedEmail
                        }
                    } catch { }
                }

                if ($emailForPermanentDelete) {
                    $emailForPermanentDelete.Delete()
                    Write-Host "`n[OK] EMAIL PERMANENTLY DELETED" -ForegroundColor Green
                } else {
                    Write-Host "`n[SAFE-FAIL] Permanent delete was not completed because message identity could not be verified." -ForegroundColor Yellow
                    Write-Host "Message was not permanently purged to avoid deleting an ambiguous item." -ForegroundColor Gray
                }
            } else {
                $targetEmail.Delete()
                Write-Host "`n[OK] EMAIL DELETED (moved to Deleted Items)" -ForegroundColor Green
            }

            Write-Host "Subject: $subject" -ForegroundColor Yellow
            Write-Host "From: $senderName <$senderAddr>"
            Write-Host "Date: $date"
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
