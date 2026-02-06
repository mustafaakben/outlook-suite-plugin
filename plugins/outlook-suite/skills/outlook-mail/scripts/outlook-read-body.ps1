param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [int]$Days = 7
)

# Outlook Email Body Reader
# Usage: .\outlook-read-body.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-read-body.ps1 -Index 3 -Days 2 (fallback)

$Outlook = $null
try {
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")
    $email = $null

    if ($EntryID) {
        # EntryID = primary (direct lookup)
        $email = $Namespace.GetItemFromID($EntryID)
        if (-not $email) {
            Write-Host "ERROR: Email not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-find.ps1 to get valid EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
        # Index = fallback (scan)
        $Inbox = $Namespace.GetDefaultFolder(6)
        $since = (Get-Date).AddDays(-$Days).ToString("g")
        $filter = "[ReceivedTime] >= '$since'"
        $items = $Inbox.Items.Restrict($filter)
        $items.Sort("[ReceivedTime]", $true)

        $count = 0
        foreach ($item in $items) {
            $count++
            if ($count -eq $Index) {
                $email = $item
                break
            }
        }

        if (-not $email) {
            Write-Host "Email at index $Index not found in last $Days days." -ForegroundColor Red
            Write-Host "Tip: Use outlook-read.ps1 -Days $Days to see available emails." -ForegroundColor Gray
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target an email." -ForegroundColor Red
        Write-Host "Usage: .\outlook-read-body.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
        Write-Host "       .\outlook-read-body.ps1 -Index 3 -Days 2" -ForegroundColor Gray
    }

    if ($email) {
        $senderAddr = $email.SenderEmailAddress
        if ($senderAddr -match "/O=") { $senderAddr = $email.SenderName }

        Write-Host "`n=== EMAIL ===" -ForegroundColor Cyan
        Write-Host "Subject: $($email.Subject)" -ForegroundColor Yellow
        Write-Host "From: $($email.SenderName) <$senderAddr>"
        Write-Host "Date: $($email.ReceivedTime.ToString('g'))"
        Write-Host "To: $($email.To)"
        if ($email.CC) { Write-Host "CC: $($email.CC)" }
        if ($email.BCC) { Write-Host "BCC: $($email.BCC)" }
        Write-Host "`n--- BODY ---" -ForegroundColor Gray
        Write-Host $email.Body
        Write-Host "--- END ---" -ForegroundColor Gray

        if ($email.Attachments.Count -gt 0) {
            Write-Host "`nAttachments:" -ForegroundColor Yellow
            foreach ($att in $email.Attachments) {
                Write-Host "  - $($att.FileName)"
            }
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
