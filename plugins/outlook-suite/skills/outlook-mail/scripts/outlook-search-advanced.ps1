param(
    [string]$Subject = "",
    [string]$From = "",
    [string]$To = "",
    [string]$Body = "",
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [ValidateSet("", "Low", "Normal", "High")]
    [string]$Importance = "",
    [switch]$HasAttachment,
    [switch]$Flagged,
    [switch]$Unread,
    [string]$Category = "",
    [string]$Folder = "Inbox",
    [int]$Limit = 20
)

# Outlook Advanced Search Script
# Usage: .\outlook-search-advanced.ps1 -HasAttachment -Unread -Limit 10
# Date range: .\outlook-search-advanced.ps1 -DateFrom "2026-01-01" -DateTo "2026-01-31"

$importanceMap = @{
    "Low" = 0
    "Normal" = 1
    "High" = 2
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
    } elseif ($Folder -eq "Sent") {
        $targetFolder = $Namespace.GetDefaultFolder(5)
    } elseif ($Folder -eq "Drafts") {
        $targetFolder = $Namespace.GetDefaultFolder(16)
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
        $targetFolder = $Namespace.GetDefaultFolder(6)
        Write-Host "Folder '$Folder' not found, using Inbox." -ForegroundColor Yellow
    }

    # Adjust DateTo to end of day (user expects full day included)
    if ($DateTo) {
        $DateTo = $DateTo.Date.AddDays(1).AddSeconds(-1)
    }

    Write-Host "`n=== ADVANCED SEARCH ===" -ForegroundColor Cyan
    Write-Host "Searching in: $($targetFolder.Name)" -ForegroundColor Gray

    # Build filter criteria display (only show active filters)
    $criteria = @()
    if ($Subject) { $criteria += "Subject: '$Subject'" }
    if ($From) { $criteria += "From: '$From'" }
    if ($To) { $criteria += "To: '$To'" }
    if ($Body) { $criteria += "Body: '$Body'" }
    if ($DateFrom) { $criteria += "From date: $($DateFrom.ToString('yyyy-MM-dd'))" }
    if ($DateTo) { $criteria += "To date: $($DateTo.ToString('yyyy-MM-dd'))" }
    if ($Importance) { $criteria += "Importance: $Importance" }
    if ($HasAttachment) { $criteria += "Has attachment" }
    if ($Flagged) { $criteria += "Flagged" }
    if ($Unread) { $criteria += "Unread" }
    if ($Category) { $criteria += "Category: '$Category'" }

    if ($criteria.Count -gt 0) {
        Write-Host "Filters: $($criteria -join ' | ')" -ForegroundColor Gray
    }
    Write-Host ""

    # Build JET date pre-filter for server-side narrowing
    $filterParts = @()
    if ($DateFrom) {
        $fromStr = $DateFrom.ToString("g")
        $filterParts += "[ReceivedTime] >= '$fromStr'"
    }
    if ($DateTo) {
        $toStr = $DateTo.ToString("g")
        $filterParts += "[ReceivedTime] <= '$toStr'"
    }

    $items = $targetFolder.Items
    if ($filterParts.Count -gt 0) {
        $jetFilter = $filterParts -join " AND "
        $items = $items.Restrict($jetFilter)
    }
    $items.Sort("[ReceivedTime]", $true)

    $count = 0
    $scanned = 0
    $maxScan = 1000

    foreach ($email in $items) {
        $scanned++
        if ($scanned -gt $maxScan) { break }

        $senderName = [string]$email.SenderName
        $senderAddr = [string]$email.SenderEmailAddress
        if ($senderAddr -match "/O=") { $senderAddr = $senderName }

        # Client-side text filters
        if ($Subject -and $email.Subject -notlike "*$Subject*") { continue }
        if ($From -and $senderName -notlike "*$From*" -and $senderAddr -notlike "*$From*") { continue }
        if ($To -and $email.To -notlike "*$To*") { continue }
        if ($Body -and $email.Body -notlike "*$Body*") { continue }

        # Client-side property filters
        if ($Importance -and $email.Importance -ne $importanceMap[$Importance]) { continue }
        if ($HasAttachment -and $email.Attachments.Count -eq 0) { continue }
        if ($Flagged -and $email.FlagStatus -ne 2) { continue }
        if ($Unread -and -not $email.UnRead) { continue }
        if ($Category -and $email.Categories -notlike "*$Category*") { continue }

        $count++

        # Display inline (no two-pass accumulation)
        $markers = ""
        if ($email.UnRead) { $markers += "[*]" }
        if ($email.Attachments.Count -gt 0) { $markers += "[A]" }
        if ($email.FlagStatus -eq 2) { $markers += "[F]" }
        if ($email.Importance -eq 2) { $markers += "[!]" }

        Write-Host "$count. $markers $($email.Subject)" -ForegroundColor Yellow
        Write-Host "      From: $senderName <$senderAddr>"
        Write-Host "      Date: $($email.ReceivedTime)"
        Write-Host "      EntryID: $($email.EntryID)" -ForegroundColor Gray
        if ($email.Categories) {
            Write-Host "      Categories: $($email.Categories)" -ForegroundColor Gray
        }
        Write-Host ""

        if ($count -ge $Limit) { break }
    }

    if ($count -eq 0) {
        Write-Host "No matching emails found." -ForegroundColor Gray
        Write-Host "Tip: Try broadening your filters or increasing -Limit." -ForegroundColor Gray
    } else {
        Write-Host "--- Found: $count email(s) ---" -ForegroundColor Cyan
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
