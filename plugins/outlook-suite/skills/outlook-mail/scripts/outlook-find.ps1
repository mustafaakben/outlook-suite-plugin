param(
    [int]$Days = 7,
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [string]$From = "",
    [string]$To = "",
    [string]$Subject = "",
    [string]$BodyContains = "",
    [switch]$UnreadOnly,
    [switch]$ReadOnly,
    [switch]$HasAttachment,
    [switch]$Flagged,
    [ValidateSet("", "Low", "Normal", "High")]
    [string]$Importance = "",
    [string]$Category = "",
    [string]$Folder = "Inbox",
    [int]$Limit = 50,
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Find Script — EntryID retrieval with search/filter conditions
# Usage: .\outlook-find.ps1 -Days 1 -UnreadOnly
# Usage: .\outlook-find.ps1 -Subject "lunch" -Days 1
# Usage: .\outlook-find.ps1 -BodyContains "invoice" -Days 30

$importanceMap = @{
    "Low"    = 0
    "Normal" = 1
    "High"   = 2
}

$importanceNames = @{
    0 = "Low"
    1 = "Normal"
    2 = "High"
}

$Outlook = $null
try {
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")

    # Get folder
    $targetFolder = $null
    if ($Folder -eq "Inbox") {
        $targetFolder = $Namespace.GetDefaultFolder(6)
    } elseif ($Folder -eq "Sent Items" -or $Folder -eq "Sent") {
        $targetFolder = $Namespace.GetDefaultFolder(5)
    } elseif ($Folder -eq "Drafts") {
        $targetFolder = $Namespace.GetDefaultFolder(16)
    } elseif ($Folder -eq "Deleted Items") {
        $targetFolder = $Namespace.GetDefaultFolder(3)
    } elseif ($Folder -eq "Junk Email" -or $Folder -eq "Junk") {
        $targetFolder = $Namespace.GetDefaultFolder(23)
    } elseif ($Folder -eq "Archive") {
        # Archive folder - search by name
        function Find-OutlookFolder($parent, $name) {
            foreach ($f in $parent.Folders) {
                if ($f.Name -eq $name) { return $f }
                $sub = Find-OutlookFolder $f $name
                if ($sub) { return $sub }
            }
            return $null
        }
        foreach ($account in $Namespace.Folders) {
            $targetFolder = Find-OutlookFolder $account "Archive"
            if ($targetFolder) { break }
        }
    } else {
        # Recursive search for folder by name
        function Find-OutlookFolder($parent, $name) {
            foreach ($f in $parent.Folders) {
                if ($f.Name -eq $name) { return $f }
                $sub = Find-OutlookFolder $f $name
                if ($sub) { return $sub }
            }
            return $null
        }
        foreach ($account in $Namespace.Folders) {
            $targetFolder = Find-OutlookFolder $account $Folder
            if ($targetFolder) { break }
        }
    }

    if (-not $targetFolder) {
        Write-Host "Folder '$Folder' not found, using Inbox." -ForegroundColor Yellow
        $targetFolder = $Namespace.GetDefaultFolder(6)
    }

    # Build date range
    if ($DateFrom -and $DateTo) {
        # Explicit date range overrides -Days
        $DateTo = $DateTo.Date.AddDays(1).AddSeconds(-1)
    } elseif ($DateFrom -and -not $DateTo) {
        # DateFrom without DateTo: from that date to now
        $DateTo = Get-Date
    } elseif (-not $DateFrom -and $DateTo) {
        # DateTo without DateFrom: use -Days before DateTo
        $DateTo = $DateTo.Date.AddDays(1).AddSeconds(-1)
        $DateFrom = $DateTo.AddDays(-$Days)
    } else {
        # Default: use -Days from now
        $DateFrom = (Get-Date).AddDays(-$Days)
    }

    # Build JET date pre-filter (server-side)
    $filterParts = @()
    $fromStr = $DateFrom.ToString("g")
    $filterParts += "[ReceivedTime] >= '$fromStr'"
    if ($DateTo) {
        $toStr = $DateTo.ToString("g")
        $filterParts += "[ReceivedTime] <= '$toStr'"
    }

    # UnreadOnly can be server-side filtered via JET
    if ($UnreadOnly) {
        $filterParts += "[UnRead] = True"
    }
    if ($ReadOnly) {
        $filterParts += "[UnRead] = False"
    }

    $jetFilter = $filterParts -join " AND "

    # Display search info
    Write-Host "`n=== OUTLOOK FIND ===" -ForegroundColor Cyan
    Write-Host "Folder: $($targetFolder.Name)" -ForegroundColor Gray

    $criteria = @()
    if ($DateFrom) { $criteria += "From: $($DateFrom.ToString('g'))" }
    if ($DateTo) { $criteria += "To: $($DateTo.ToString('g'))" }
    if ($From) { $criteria += "Sender: '$From'" }
    if ($To) { $criteria += "Recipient: '$To'" }
    if ($Subject) { $criteria += "Subject: '$Subject'" }
    if ($BodyContains) { $criteria += "Body: '$BodyContains'" }
    if ($UnreadOnly) { $criteria += "Unread only" }
    if ($ReadOnly) { $criteria += "Read only" }
    if ($HasAttachment) { $criteria += "Has attachments" }
    if ($Flagged) { $criteria += "Flagged" }
    if ($Importance) { $criteria += "Importance: $Importance" }
    if ($Category) { $criteria += "Category: '$Category'" }

    if ($criteria.Count -gt 0) {
        Write-Host "Filters: $($criteria -join ' | ')" -ForegroundColor Gray
    }
    Write-Host ""

    # Execute search
    $items = $targetFolder.Items.Restrict($jetFilter)
    $items.Sort("[ReceivedTime]", $true)

    $count = 0
    $scanned = 0
    $maxScan = 2000

    foreach ($email in $items) {
        $scanned++
        if ($scanned -gt $maxScan) { break }

        # Client-side text filters (accuracy > performance)
        if ($Subject -and $email.Subject -notlike "*$Subject*") { continue }
        if ($From) {
            $senderMatch = $false
            if ($email.SenderName -like "*$From*") { $senderMatch = $true }
            if ($email.SenderEmailAddress -like "*$From*") { $senderMatch = $true }
            if (-not $senderMatch) { continue }
        }
        if ($To -and $email.To -notlike "*$To*") { continue }
        if ($BodyContains -and $email.Body -notlike "*$BodyContains*") { continue }

        # Client-side property filters
        if ($HasAttachment -and $email.Attachments.Count -eq 0) { continue }
        if ($Flagged -and $email.FlagStatus -ne 2) { continue }
        if ($Importance -and $email.Importance -ne $importanceMap[$Importance]) { continue }
        if ($Category -and $email.Categories -notlike "*$Category*") { continue }

        $count++

        # Sender info
        $senderAddr = $email.SenderEmailAddress
        if ($senderAddr -match "/O=") { $senderAddr = $email.SenderName }

        # Status
        $status = if ($email.UnRead) { "[UNREAD]" } else { "[READ]" }
        $statusColor = if ($email.UnRead) { "Yellow" } else { "Gray" }

        # Body snippet — plain text, first ~150 chars, trimmed
        $snippet = ""
        try {
            $bodyText = $email.Body
            if ($bodyText) {
                $bodyText = $bodyText -replace '<[^>]+>', ''
                $bodyText = $bodyText -replace '\r?\n', ' '
                $bodyText = $bodyText -replace '\s+', ' '
                $bodyText = $bodyText.Trim()
                if ($bodyText.Length -gt 150) {
                    $snippet = $bodyText.Substring(0, 150) + "..."
                } else {
                    $snippet = $bodyText
                }
            }
        } catch { }

        # Attachments info
        $attachInfo = ""
        if ($email.Attachments.Count -gt 0) {
            $attachNames = @()
            for ($i = 1; $i -le $email.Attachments.Count; $i++) {
                $attachNames += $email.Attachments.Item($i).FileName
            }
            $attachInfo = "$($email.Attachments.Count) attachment(s): $($attachNames -join ', ')"
        }

        # Display result
        Write-Host "--- #$count ---" -ForegroundColor Cyan
        Write-Host "  EntryID: $($email.EntryID)"
        Write-Host "  Subject: $($email.Subject)" -ForegroundColor Yellow
        Write-Host "  From: $($email.SenderName) <$senderAddr>"
        Write-Host "  Date: $($email.ReceivedTime.ToString('g'))"
        Write-Host "  Status: $status" -ForegroundColor $statusColor

        if ($snippet) {
            if ($StripLinks) { $snippet = Strip-Links $snippet }
            Write-Host "  Body: $snippet" -ForegroundColor Gray
        }

        if ($attachInfo) {
            Write-Host "  Attachments: $attachInfo" -ForegroundColor Gray
        }

        # Conditional fields (only show when non-default)
        if ($email.Importance -ne 1) {
            $impName = $importanceNames[[int]$email.Importance]
            Write-Host "  Importance: $impName" -ForegroundColor Magenta
        }
        if ($email.FlagStatus -eq 2) {
            Write-Host "  Flag: Flagged" -ForegroundColor Red
        }
        if ($email.Categories) {
            Write-Host "  Categories: $($email.Categories)" -ForegroundColor Green
        }

        Write-Host ""

        if ($count -ge $Limit) {
            Write-Host "(Showing first $Limit results. Use -Limit to increase.)" -ForegroundColor Gray
            break
        }
    }

    if ($count -eq 0) {
        Write-Host "No matching emails found." -ForegroundColor Gray
        Write-Host "Tips:" -ForegroundColor Gray
        Write-Host "  - Try increasing -Days (current: $Days)" -ForegroundColor Gray
        Write-Host "  - Try broadening your filters" -ForegroundColor Gray
        Write-Host "  - Check the folder name with outlook-folders.ps1" -ForegroundColor Gray
    } else {
        Write-Host "=== Found: $count email(s) ===" -ForegroundColor Cyan
        Write-Host "Use the EntryID with action scripts, e.g.:" -ForegroundColor Gray
        Write-Host "  outlook-read-body.ps1 -EntryID ""<EntryID>""" -ForegroundColor Gray
        Write-Host "  outlook-reply.ps1 -EntryID ""<EntryID>"" -Body ""message""" -ForegroundColor Gray
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
