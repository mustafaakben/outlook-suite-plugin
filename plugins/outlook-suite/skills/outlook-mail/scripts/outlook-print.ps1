param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [int]$Days = 7,
    [switch]$ConfirmPrint
)

# Outlook Print Email Script
# Usage: .\outlook-print.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-print.ps1 -Index 1 (fallback)
# Print confirmation: .\outlook-print.ps1 -Index 1 -ConfirmPrint

$Outlook = $null
$Namespace = $null

try {
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
        Write-Host "Usage: .\outlook-print.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Get sender email with Exchange fallback
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

        Write-Host "`n=== PRINTING EMAIL ===" -ForegroundColor Cyan
        Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
        Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
        Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))"
        Write-Host ""

        if (-not $ConfirmPrint) {
            Write-Host "[Preview only] Add -ConfirmPrint to send this email to the default printer." -ForegroundColor Cyan
        } else {
            # Print the email
            $targetEmail.PrintOut()

            Write-Host "[OK] Email sent to default printer." -ForegroundColor Green
            Write-Host ""
            Write-Host "Note: The print job has been submitted to your default printer." -ForegroundColor Gray
            Write-Host "Check your printer queue if the document doesn't print immediately." -ForegroundColor Gray
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
