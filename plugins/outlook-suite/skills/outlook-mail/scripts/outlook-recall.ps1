param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [switch]$DeleteUnread,
    [switch]$ConfirmDeleteFromSent,
    [int]$Days = 7
)

# Outlook Recall Email Script
# Attempt to recall a sent email (Exchange only)
# Usage: .\outlook-recall.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-recall.ps1 -Index 1 (fallback)
# Delete from Sent Items only: .\outlook-recall.ps1 -EntryID "00000000..." -DeleteUnread -ConfirmDeleteFromSent

# IMPORTANT NOTES:
# - Recall only works with Microsoft Exchange/Office 365
# - Both sender and recipient must be on the same Exchange server
# - Recall fails if the recipient has already read the email
# - Recall fails if the email was moved to a different folder
# - Recall fails if the recipient is using a non-Exchange email client
# - Even when "successful", recall notifications may reveal the original message

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
            Write-Host "Tip: Use outlook-find.ps1 -Folder ""Sent Items"" to get valid EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
        $SentItems = $Namespace.GetDefaultFolder(5)  # olFolderSentMail = 5

        # Get sent emails with locale-safe date
        $since = (Get-Date).AddDays(-$Days).ToString("g")
        $filter = "[SentOn] >= '$since'"

        $items = $SentItems.Items.Restrict($filter)
        $items.Sort("[SentOn]", $true)

        # Find the email at the specified index
        $count = 0
        foreach ($email in $items) {
            $count++
            if ($count -eq $Index) {
                $targetEmail = $email
                break
            }
        }

        if (-not $targetEmail) {
            Write-Host "`nSent email at index $Index not found in last $Days days." -ForegroundColor Red
            Write-Host "Note: Index is based on Sent Items folder." -ForegroundColor Gray
            Write-Host "Tip: Try increasing -Days if the email is older." -ForegroundColor Gray
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target an email." -ForegroundColor Red
        Write-Host "Usage: .\outlook-recall.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if ($targetEmail) {
        Write-Host "`n=== RECALL EMAIL ===" -ForegroundColor Cyan
        Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
        Write-Host "To: $($targetEmail.To)"
        Write-Host "Sent: $($targetEmail.SentOn.ToString('g'))" -ForegroundColor Gray

        Write-Host "`n--- Warning ---" -ForegroundColor Yellow
        Write-Host "Email recall has significant limitations:" -ForegroundColor Yellow
        Write-Host "  - Only works within Exchange/Office 365" -ForegroundColor Gray
        Write-Host "  - Fails if recipient already read the email" -ForegroundColor Gray
        Write-Host "  - Fails if email moved from Inbox" -ForegroundColor Gray
        Write-Host "  - May not work for external recipients" -ForegroundColor Gray

        # Attempt recall
        try {
            if ($DeleteUnread) {
                Write-Host "`nRequested action: delete this message from your Sent Items only." -ForegroundColor Yellow
                Write-Host "This is not a recipient recall." -ForegroundColor Yellow

                if (-not $ConfirmDeleteFromSent) {
                    Write-Host "`n[!] Add -ConfirmDeleteFromSent to proceed with Sent Items deletion." -ForegroundColor Cyan
                    Write-Host "No changes were made." -ForegroundColor Gray
                } else {
                    $targetEmail.Delete()
                    Write-Host "`nAction complete: removed from your Sent Items." -ForegroundColor Green
                    Write-Host "No recipient-side recall was sent." -ForegroundColor Gray
                }
            } else {
                # Standard recall - COM doesn't have direct Recall method
                Write-Host "`nAttempting to initiate recall..." -ForegroundColor Cyan
                Write-Host "`nLimitation: Programmatic recall is not fully supported via COM." -ForegroundColor Yellow
                Write-Host "To recall this email:" -ForegroundColor Cyan
                Write-Host "  1. Open Outlook" -ForegroundColor Gray
                Write-Host "  2. Go to Sent Items" -ForegroundColor Gray
                Write-Host "  3. Open the email" -ForegroundColor Gray
                Write-Host "  4. Click Message > Actions > Recall This Message" -ForegroundColor Gray
                Write-Host "`n--- Recall Status ---" -ForegroundColor Cyan
                Write-Host "No recall request was sent automatically by this script." -ForegroundColor Gray
                Write-Host "Use the Outlook steps above to attempt recipient recall." -ForegroundColor Gray
                Write-Host "Recipients may still see the original message." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "`nOperation failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "If this was a recall attempt, recall may not be available for this message." -ForegroundColor Gray
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
