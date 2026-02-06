param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [string]$To,

    [string]$Body = "",
    [string]$CC = "",
    [string]$BCC = "",
    [int]$Days = 7,
    [switch]$Confirm
)

# Outlook Forward Script
# Usage: .\outlook-forward.ps1 -EntryID "00000000..." -To "colleague@example.com" (preferred)
# Usage: .\outlook-forward.ps1 -Index 1 -To "colleague@example.com" (fallback)
# Without -Confirm, creates a draft in the Drafts folder.

$Outlook = $null
$forward = $null

try {
    # Connect to Outlook (fast path if already running)
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
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target an email." -ForegroundColor Red
        Write-Host "Usage: .\outlook-forward.ps1 -EntryID ""00000000..."" -To ""email@example.com""" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Sender email with Exchange fallback
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "/O=") { $senderAddr = $targetEmail.SenderName }

        # Create forward
        $forward = $targetEmail.Forward()
        $forward.To = $To
        if ($CC) { $forward.CC = $CC }
        if ($BCC) { $forward.BCC = $BCC }

        # HTML-aware body insertion
        if ($Body) {
            Add-Type -AssemblyName System.Web
            if ($forward.BodyFormat -eq 2) {
                # HTML format - wrap in HTML and prepend
                $htmlBody = "<div style='font-family:Calibri,sans-serif;font-size:11pt;'>$([System.Web.HttpUtility]::HtmlEncode($Body).Replace("`n","<br>"))</div><br>"
                $forward.HTMLBody = $htmlBody + $forward.HTMLBody
            } else {
                # Plain text
                $forward.Body = $Body + "`n`n" + $forward.Body
            }
        }

        # Body preview for display
        $bodyPreview = $Body
        if ($bodyPreview.Length -gt 200) {
            $bodyPreview = $bodyPreview.Substring(0, 200) + "..."
        }

        if ($Confirm) {
            # Show confirmation summary before sending
            Write-Host "`n=== FORWARD CONFIRMATION ===" -ForegroundColor Yellow
            Write-Host "Original: $($targetEmail.Subject)" -ForegroundColor Yellow
            Write-Host "Original From: $($targetEmail.SenderName) <$senderAddr>"
            Write-Host "Original Date: $($targetEmail.ReceivedTime.ToString("g"))"
            Write-Host "Forwarding To: $To" -ForegroundColor Cyan
            if ($CC) {
                Write-Host "CC: $CC" -ForegroundColor Gray
            }
            if ($BCC) {
                Write-Host "BCC: $BCC" -ForegroundColor Gray
            }
            if ($Body) {
                Write-Host "`nYour message:" -ForegroundColor Cyan
                Write-Host $bodyPreview
            }

            $forward.Send()
            Write-Host "`n=== EMAIL FORWARDED ===" -ForegroundColor Green
        } else {
            # Save as draft
            $forward.Save()

            Write-Host "`n=== FORWARD DRAFT CREATED ===" -ForegroundColor Green
            Write-Host "Original: $($targetEmail.Subject)" -ForegroundColor Yellow
            Write-Host "Original From: $($targetEmail.SenderName) <$senderAddr>"
            Write-Host "Original Date: $($targetEmail.ReceivedTime.ToString("g"))"
            Write-Host "Forwarding To: $To" -ForegroundColor Cyan
            if ($CC) {
                Write-Host "CC: $CC" -ForegroundColor Gray
            }
            if ($BCC) {
                Write-Host "BCC: $BCC" -ForegroundColor Gray
            }
            if ($Body) {
                Write-Host "`nYour message:" -ForegroundColor Cyan
                Write-Host $bodyPreview
            }
            Write-Host "`nDraft saved. Check your Drafts folder." -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($forward) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($forward) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
