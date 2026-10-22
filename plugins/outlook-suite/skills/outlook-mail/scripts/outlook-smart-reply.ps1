param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [ValidateRange(0, 3)]
    [int]$Use = 0,                   # 1-3: save a reply draft with suggestion N
    [switch]$ReplyAll,
    [switch]$RemoveAutoSignature,
    [int]$Days = 2,
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [switch]$UnreadOnly,
    [string]$From = "",
    [string]$Subject = "",
    [string]$Folder = "Inbox",
    [int]$Limit = 20,
    [int]$MaxScan = 300,
    [switch]$Json
)

# Outlook Smart Reply: Copilot's suggested replies (the quick-reply chips in Outlook).
# Exchange stores up to 3 per email; about a quarter of mail has them.
#   One email:   -EntryID <id>             its suggestions, numbered 1-3 in Copilot's order
#   List:        (no -EntryID)             emails that have suggestions, with them
#   Draft:       -EntryID <id> -Use 2      reply draft in the thread with suggestion 2 (never sent;
#                                          send later with outlook-send-draft.ps1 -EntryID <draft> -Confirm)
#
# Usage: .\outlook-smart-reply.ps1 -Days 2 -UnreadOnly
# Usage: .\outlook-smart-reply.ps1 -EntryID "00000000..."
# Usage: .\outlook-smart-reply.ps1 -EntryID "00000000..." -Use 1 -ReplyAll

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    if ($Use -gt 0) {
        if (-not $EntryID) { Stop-WithError "-Use needs -EntryID (the email to reply to)." }
        $reply = @{ EntryID = $EntryID; SmartReply = $Use; ReplyAll = $ReplyAll; RemoveAutoSignature = $RemoveAutoSignature; Json = $Json }
        if ($StoreID) { $reply.StoreID = $StoreID }
        & (Join-Path $PSScriptRoot "outlook-reply.ps1") @reply
        exit $LASTEXITCODE
    }

    $null = Connect-Outlook

    if ($EntryID) {
        $item = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID
        $ins = Get-CopilotInsight $item
        $sugg = @($ins.SmartReplies)
        if ($Json) {
            Write-Json ([ordered]@{ EntryID = $item.EntryID; Subject = $item.Subject; From = $item.SenderName; Date = Format-MailDateIso $item.ReceivedTime; SmartReplies = $sugg })
            exit 0
        }
        Write-Host "=== SMART REPLIES ===" -ForegroundColor Cyan
        Write-Host "  EntryID: $($item.EntryID)"
        Write-Host "  Subject: $($item.Subject)" -ForegroundColor Yellow
        Write-Host "  From: $($item.SenderName)"
        if ($sugg.Count -eq 0) {
            Write-Host "  Copilot has no suggested replies for this email." -ForegroundColor Gray
        } else {
            for ($i = 0; $i -lt $sugg.Count; $i++) { Write-Host "  [$($i + 1)] $($sugg[$i])" -ForegroundColor Green }
            Write-Host "Next: outlook-smart-reply.ps1 -EntryID <EntryID> -Use <1-$($sugg.Count)> [-ReplyAll]" -ForegroundColor DarkGray
        }
        exit 0
    }

    # List: candidates from one table query, then suggestions read per email (about 7 ms each).
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
    $range = Resolve-DateRange -Days $Days -DateFrom $DateFrom -DateTo $DateTo
    $query = @{ Folder = $target; From = $range[0]; Subject = $Subject; Sender = $From; UnreadOnly = $UnreadOnly; Limit = $MaxScan }
    if ($range[1]) { $query.To = $range[1] }
    $result = Find-MailRecords @query
    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($r in $result.Records) {
        if ($hits.Count -ge $Limit) { break }
        try { $item = $Namespace.GetItemFromID($r.EntryID, $r.StoreID) } catch { continue }
        $sugg = @((Get-CopilotInsight $item).SmartReplies)
        if ($sugg.Count -eq 0) { continue }
        $hits.Add([ordered]@{ EntryID = $r.EntryID; Date = $r.Date; From = $r.From; Subject = $r.Subject; Unread = $r.Unread; SmartReplies = $sugg })
    }
    $scanned = $result.Records.Count

    if ($Json) {
        $items = @(foreach ($h in $hits) { ConvertTo-JsonRecord $h })
        Write-Json ([ordered]@{ folder = $target.FolderPath; scanned = $scanned; matched = $result.Total; count = $items.Count; items = $items })
        exit 0
    }
    Write-Host "=== SMART REPLIES ===" -ForegroundColor Cyan
    Write-Host "Folder: $($target.FolderPath)"
    Write-Host "Filters: since $(Format-MailDate $range[0])$(if ($UnreadOnly) { ' | unread' })$(if ($From) { " | from: '$From'" })$(if ($Subject) { " | subject: '$Subject'" })" -ForegroundColor Gray
    Write-Host ""
    $n = 0
    foreach ($h in $hits) {
        $n++
        Write-Host "[$n] $(Format-MailDate $h.Date)  $(if ($h.Unread) { 'UNREAD' } else { 'read' })" -ForegroundColor Cyan
        Write-Host "  Subject: $($h.Subject)" -ForegroundColor Yellow
        Write-Host "  From: $($h.From)"
        for ($i = 0; $i -lt $h.SmartReplies.Count; $i++) { Write-Host "  [$($i + 1)] $($h.SmartReplies[$i])" -ForegroundColor Green }
        Write-Host "  EntryID: $($h.EntryID)" -ForegroundColor DarkGray
        Write-Host ""
    }
    if ($n -eq 0) { Write-Host "No email in this range has suggested replies. Try a larger -Days." -ForegroundColor Gray }
    $more = if ($scanned -lt $result.Total) { " (checked the newest $scanned of $($result.Total); raise -MaxScan)" } else { "" }
    Write-Host "=== $n with suggestions, $scanned checked$more ===" -ForegroundColor Cyan
    Write-Host "Next: outlook-smart-reply.ps1 -EntryID <EntryID> -Use <1-3> [-ReplyAll]" -ForegroundColor DarkGray
}
catch {
    Stop-WithError $_
}
