param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [string]$Body,

    [int]$Days = 7,
    [switch]$ReplyAll,
    [switch]$Confirm
)

# Outlook Reply Script
# Usage: .\outlook-reply.ps1 -EntryID "00000000..." -Body "Thank you for your email." (preferred)
# Usage: .\outlook-reply.ps1 -Index 1 -Body "Thank you for your email." (fallback)
# Without -Confirm, creates a draft in the Drafts folder.

$Outlook = $null
$reply = $null

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
        # EntryID = primary (direct lookup)
        $targetEmail = $Namespace.GetItemFromID($EntryID)
        if (-not $targetEmail) {
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
        Write-Host "Usage: .\outlook-reply.ps1 -EntryID ""00000000..."" -Body ""message""" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Sender email with Exchange fallback
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "/O=") { $senderAddr = $targetEmail.SenderName }

        # Create reply
        if ($ReplyAll) {
            $reply = $targetEmail.ReplyAll()
            $replyType = "Reply All"
        } else {
            $reply = $targetEmail.Reply()
            $replyType = "Reply"
        }

        # HTML-aware body insertion
        Add-Type -AssemblyName System.Web
        if ($reply.BodyFormat -eq 2) {
            # HTML format - wrap reply body in HTML and prepend
            $htmlBody = "<div style='font-family:Calibri,sans-serif;font-size:11pt;'>$([System.Web.HttpUtility]::HtmlEncode($Body).Replace("`n","<br>"))</div><br>"
            $reply.HTMLBody = $htmlBody + $reply.HTMLBody
        } else {
            # Plain text
            $reply.Body = $Body + "`n`n" + $reply.Body
        }

        # Body preview for display
        $bodyPreview = $Body
        if ($bodyPreview.Length -gt 200) {
            $bodyPreview = $bodyPreview.Substring(0, 200) + "..."
        }

        if ($Confirm) {
            # Show confirmation summary before sending
            Write-Host "`n=== $replyType CONFIRMATION ===" -ForegroundColor Yellow
            Write-Host "Original: $($targetEmail.Subject)" -ForegroundColor Yellow
            Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
            Write-Host "To: $($reply.To)" -ForegroundColor Gray
            if ($reply.CC) {
                Write-Host "CC: $($reply.CC)" -ForegroundColor Gray
            }
            if ($reply.BCC) {
                Write-Host "BCC: $($reply.BCC)" -ForegroundColor Gray
            }
            Write-Host "`nYour reply:" -ForegroundColor Cyan
            Write-Host $bodyPreview

            $reply.Send()
            Write-Host "`n=== $replyType SENT ===" -ForegroundColor Green
        } else {
            # Save as draft
            $reply.Save()

            Write-Host "`n=== $replyType DRAFT CREATED ===" -ForegroundColor Green
            Write-Host "Original: $($targetEmail.Subject)" -ForegroundColor Yellow
            Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
            Write-Host "To: $($reply.To)" -ForegroundColor Gray
            if ($reply.CC) {
                Write-Host "CC: $($reply.CC)" -ForegroundColor Gray
            }
            if ($reply.BCC) {
                Write-Host "BCC: $($reply.BCC)" -ForegroundColor Gray
            }
            Write-Host "`nYour reply:" -ForegroundColor Cyan
            Write-Host $bodyPreview
            Write-Host "`nDraft saved. Check your Drafts folder." -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($reply) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($reply) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
