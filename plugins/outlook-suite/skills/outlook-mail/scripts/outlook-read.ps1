param(
    [int]$Days = 1,
    [int]$Limit = 20,
    [switch]$UnreadOnly
)

# Outlook Email Reader
# Usage: .\outlook-read.ps1 -Days 2 -Limit 50 -UnreadOnly

$Outlook = $null
try {
    $Outlook = New-Object -ComObject Outlook.Application
    Start-Sleep -Milliseconds 500
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Inbox = $Namespace.GetDefaultFolder(6)

    $since = (Get-Date).AddDays(-$Days).ToString("g")

    if ($UnreadOnly) {
        $filter = "[ReceivedTime] >= '$since' AND [UnRead] = True"
    } else {
        $filter = "[ReceivedTime] >= '$since'"
    }

    $items = $Inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)

    Write-Host "`n=== EMAILS (Last $Days days) ===" -ForegroundColor Cyan
    if ($UnreadOnly) { Write-Host "(Unread only)" -ForegroundColor Gray }
    Write-Host ""

    $count = 0
    foreach ($email in $items) {
        $count++
        $unreadMark = if ($email.UnRead) { "[*]" } else { "   " }
        $senderAddr = $email.SenderEmailAddress
        if ($senderAddr -match "/O=") { $senderAddr = $email.SenderName }
        Write-Host "$count. $unreadMark $($email.Subject)" -ForegroundColor Yellow
        Write-Host "      From: $($email.SenderName) <$senderAddr>"
        Write-Host "      Date: $($email.ReceivedTime)"
        Write-Host "      EntryID: $($email.EntryID)" -ForegroundColor Gray
        Write-Host ""

        if ($count -ge $Limit) { break }
    }

    if ($count -eq 0) {
        Write-Host "No emails found in the last $Days day(s)." -ForegroundColor Gray
        Write-Host "Tip: Try increasing -Days (e.g., -Days 7)" -ForegroundColor Gray
    } else {
        Write-Host "--- Total: $count emails ---" -ForegroundColor Cyan
        Write-Host "Tip: Use outlook-read-body.ps1 -EntryID <id> (preferred) or -Index <number> to read full email content." -ForegroundColor Gray
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
