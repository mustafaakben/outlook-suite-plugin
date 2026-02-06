param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [int]$Days = 7,
    [switch]$Unread
)

# Outlook Mark Read/Unread Script
# Usage: .\outlook-mark-read.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-mark-read.ps1 -Index 1 (fallback)

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
        Write-Host "Usage: .\outlook-mark-read.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Get sender email with Exchange fallback
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

        $subject = $targetEmail.Subject
        $senderName = $targetEmail.SenderName
        $date = $targetEmail.ReceivedTime.ToString("g")
        $wasUnread = $targetEmail.UnRead

        # Set new status
        if ($Unread) {
            $newStatus = "UNREAD"
            if ($wasUnread) {
                Write-Host "`n[i] Email is already unread." -ForegroundColor Yellow
            } else {
                $targetEmail.UnRead = $true
                $targetEmail.Save()
            }
        } else {
            $newStatus = "READ"
            if (-not $wasUnread) {
                Write-Host "`n[i] Email is already read." -ForegroundColor Yellow
            } else {
                $targetEmail.UnRead = $false
                $targetEmail.Save()
            }
        }

        Write-Host "`n=== EMAIL MARKED AS $newStatus ===" -ForegroundColor Green
        Write-Host "Subject: $subject" -ForegroundColor Yellow
        Write-Host "From: $senderName <$senderAddr>"
        Write-Host "Date: $date"
        Write-Host "Previous status: $(if ($wasUnread) { 'Unread' } else { 'Read' })" -ForegroundColor Gray
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
