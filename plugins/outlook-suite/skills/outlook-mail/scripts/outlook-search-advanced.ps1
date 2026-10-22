param(
    [Alias("Folder")]
    [string]$Folders = "",
    [switch]$IncludeSubfolders,
    [switch]$WholeMailbox,
    [string]$Mailbox = "",
    [switch]$AllMailboxes,
    [string]$Subject = "",
    [string]$From = "",
    [string]$To = "",
    [Alias("Body")]
    [string]$BodyContains = "",
    [switch]$FastBody,
    [int]$Days = 30,
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [Alias("Unread")]
    [switch]$UnreadOnly,
    [switch]$HasAttachment,
    [switch]$Flagged,
    [ValidateSet("", "Low", "Normal", "High")]
    [string]$Importance = "",
    [string]$Category = "",
    [int]$Limit = 50,
    [string]$SaveAs = "",
    [switch]$Json
)

# Outlook Advanced Search: one search across many folders, subfolders or whole mailboxes,
# results merged newest first, each with its folder. Same filters as outlook-find.ps1.
#   Scope: -Folders "Inbox;Sent;Archive/2025" (default Inbox + Sent Items)
#          -IncludeSubfolders   also every folder below them
#          -WholeMailbox        every mail folder of the default mailbox (or -Mailbox)
#          -AllMailboxes        every mail folder of every mailbox (public folders excluded)
#   -SaveAs "Name"   also creates an Outlook Search Folder with these criteria (one mailbox).
# Usage: .\outlook-search-advanced.ps1 -Subject "budget" -WholeMailbox -Days 365
# Usage: .\outlook-search-advanced.ps1 -From "jane" -Folders "Inbox;Archive" -IncludeSubfolders -HasAttachment

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $range = Resolve-DateRange -Days $Days -DateFrom $DateFrom -DateTo $DateTo

    # Folders in scope.
    $scope = New-Object System.Collections.Generic.List[object]
    if ($AllMailboxes) {
        foreach ($s in (Get-SearchStores -Namespace $Namespace)) { try { foreach ($f in (Get-MailSubfolders $s.GetRootFolder())) { $scope.Add($f) } } catch { } }
    } elseif ($WholeMailbox) {
        $store = Get-RuleStore -Namespace $Namespace -Mailbox $Mailbox
        foreach ($f in (Get-MailSubfolders $store.GetRootFolder())) { $scope.Add($f) }
    } else {
        $names = if ($Folders) { @($Folders -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } else { @("Inbox", "Sent") }
        foreach ($n in $names) {
            $p = if ($Mailbox -and $n -notmatch '/') { "$Mailbox/$n" } else { $n }
            $f = $null
            try { $f = Resolve-OutlookFolder -Namespace $Namespace -Path $p } catch { if ($Mailbox) { $f = Resolve-OutlookFolder -Namespace $Namespace -Path $n } else { throw } }
            $scope.Add($f)
            if ($IncludeSubfolders) { foreach ($c in (Get-MailSubfolders $f)) { $scope.Add($c) } }
        }
    }
    $seen = @{}; $unique = @(foreach ($f in $scope) { if (-not $seen.ContainsKey($f.EntryID)) { $seen[$f.EntryID] = 1; $f } })

    $q = @{ From = $range[0]; Subject = $Subject; Sender = $From; Recipient = $To; BodyContains = $BodyContains; FastBody = $FastBody
        UnreadOnly = $UnreadOnly; HasAttachment = $HasAttachment; Flagged = $Flagged; Importance = $Importance; Category = $Category; Limit = $Limit }
    if ($range[1]) { $q.To = $range[1] }

    $all = New-Object System.Collections.Generic.List[object]
    $total = 0; $searched = 0; $skipped = 0
    $sw = [Diagnostics.Stopwatch]::StartNew()
    foreach ($f in $unique) {
        try {
            if ($f.Items.Count -eq 0) { continue }
            $res = Find-MailRecords -Folder $f @q
            $total += $res.Total; $searched++
            foreach ($r in $res.Records) { $all.Add($r) }
        } catch { $skipped++ }
    }
    $sorted = @($all | Sort-Object { [datetime]$_.Date } -Descending | Select-Object -First $Limit)
    $result = [pscustomobject]@{ Total = $total; Records = $sorted }

    $saved = ""
    if ($SaveAs) {
        $filter = (New-MailDaslFilter -From $range[0] -Subject $Subject -Sender $From -Recipient $To -UnreadOnly:$UnreadOnly -HasAttachment:$HasAttachment -Flagged:$Flagged -Importance $Importance -Category $Category -BodyIndexed $(if ($BodyContains) { $BodyContains } else { "" })) -replace '^@SQL=', ''
        $storeId = $unique[0].StoreID
        $paths = @($unique | Where-Object { $_.StoreID -eq $storeId } | Select-Object -First 30 | ForEach-Object { "'$($_.FolderPath)'" }) -join ","
        $search = $Outlook.AdvancedSearch($paths, $filter, [bool]$IncludeSubfolders, "suite-$SaveAs")
        $null = $search.Save($SaveAs)
        $saved = $SaveAs
    }

    $criteria = @("Folders searched: $searched$(if ($skipped) { " ($skipped skipped)" }) in $($sw.ElapsedMilliseconds) ms", "Since: $(Format-MailDate $range[0])")
    if ($Subject) { $criteria += "Subject: '$Subject'" }
    if ($From) { $criteria += "Sender: '$From'" }
    if ($To) { $criteria += "Recipient: '$To'" }
    if ($BodyContains) { $criteria += "Body: '$BodyContains'$(if ($FastBody) { ' (indexed)' })" }
    if ($UnreadOnly) { $criteria += "Unread only" }
    if ($HasAttachment) { $criteria += "Has attachments" }
    if ($Flagged) { $criteria += "Flagged" }
    if ($Importance) { $criteria += "Importance: $Importance" }
    if ($Category) { $criteria += "Category: '$Category'" }
    if ($saved) { $criteria += "Saved as Search Folder: '$saved'" }

    if ($Json) {
        Write-Json ([ordered]@{ total = $total; count = $sorted.Count; foldersSearched = $searched; query = $criteria; savedSearchFolder = $saved
            items = @($sorted | ForEach-Object { $o = ConvertTo-JsonRecord $_; $o }) })
        exit 0
    }
    Write-Host "=== ADVANCED SEARCH ===" -ForegroundColor Cyan
    Write-Host "Filters: $($criteria -join ' | ')" -ForegroundColor Gray
    Write-Host ""
    $n = 0
    foreach ($r in $sorted) {
        $n++
        Write-Host "--- #$n ---" -ForegroundColor Cyan
        Write-Host "  EntryID: $($r.EntryID)"
        Write-Host "  Folder: $($r.Folder)" -ForegroundColor DarkGray
        if ($r.ConversationID) { Write-Host "  ConversationID: $($r.ConversationID)" -ForegroundColor DarkGray }
        Write-Host "  Subject: $($r.Subject)" -ForegroundColor Yellow
        Write-Host "  From: $($r.From)"
        if ($r.ToLine) { Write-Host "  To: $($r.ToLine)" }
        Write-Host "  Date: $(Format-MailDate $r.Date)"
        Write-Host "  Status: $(if ($r.Unread) { 'UNREAD' } else { 'read' })"
        if ($r.Importance -and $r.Importance -ne "Normal") { Write-Host "  Importance: $($r.Importance)" -ForegroundColor Magenta }
        if ($r.Flag) { Write-Host "  Flag: $($r.Flag)" -ForegroundColor Red }
        if ($r.Categories) { Write-Host "  Categories: $($r.Categories)" -ForegroundColor Green }
        if ($r.HasAttachments) { Write-Host "  Attachments: yes" }
        if ($r.Preview) { Write-Host "  Preview: $($r.Preview)" -ForegroundColor Gray }
        Write-Host ""
    }
    if ($n -eq 0) { Write-Host "No matching emails found." -ForegroundColor Gray }
    else { Write-Host "=== Showing $n of $total matching ===" -ForegroundColor Cyan }
}
catch {
    Stop-WithError $_
}
