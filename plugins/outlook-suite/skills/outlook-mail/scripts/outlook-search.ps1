param(
    [string]$Subject = "",
    [string]$From = "",
    [string]$Body = "",
    [int]$Days = 30,
    [int]$Limit = 20
)

# Outlook Email Search
# Usage: .\outlook-search.ps1 -Subject "meeting" -From "John" -Days 7

if (-not $Subject -and -not $From -and -not $Body) {
    Write-Host "No search criteria provided. Use -Subject, -From, or -Body to filter." -ForegroundColor Red
    Write-Host "Example: .\outlook-search.ps1 -Subject `"meeting`" -From `"John`" -Days 7" -ForegroundColor Gray
    exit
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
    $Inbox = $Namespace.GetDefaultFolder(6)

    # JET filter for date (reliable server-side filtering)
    $since = (Get-Date).AddDays(-$Days).ToString("g")
    $filter = "[ReceivedTime] >= '$since'"

    $items = $Inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)

    Write-Host "`n=== SEARCH RESULTS ===" -ForegroundColor Cyan
    $activeFilters = @()
    if ($Subject) { $activeFilters += "Subject: '$Subject'" }
    if ($From) { $activeFilters += "From: '$From'" }
    if ($Body) { $activeFilters += "Body: '$Body'" }
    Write-Host "Filters: $($activeFilters -join ' | ') | Last $Days days" -ForegroundColor Gray
    Write-Host ""

    $count = 0
    foreach ($email in $items) {
        $senderName = [string]$email.SenderName
        $senderAddr = [string]$email.SenderEmailAddress
        if ($senderAddr -match "/O=") { $senderAddr = $senderName }

        # Client-side filters (accurate and reliable across all Outlook versions)
        if ($Subject -and $email.Subject -notlike "*$Subject*") { continue }
        if ($From -and $senderName -notlike "*$From*" -and $senderAddr -notlike "*$From*") { continue }
        if ($Body -and $email.Body -notlike "*$Body*") { continue }

        $count++
        $unreadMark = if ($email.UnRead) { "[*]" } else { "   " }
        Write-Host "$count. $unreadMark $($email.Subject)" -ForegroundColor Yellow
        Write-Host "      From: $senderName <$senderAddr>"
        Write-Host "      Date: $($email.ReceivedTime)"
        Write-Host "      EntryID: $($email.EntryID)" -ForegroundColor Gray
        Write-Host ""

        if ($count -ge $Limit) { break }
    }

    if ($count -eq 0) {
        Write-Host "No matching emails found." -ForegroundColor Gray
        Write-Host "Tip: Try broadening your search with -Days $($Days * 2) or fewer filters." -ForegroundColor Gray
    } else {
        Write-Host "--- Found: $count emails ---" -ForegroundColor Cyan
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
