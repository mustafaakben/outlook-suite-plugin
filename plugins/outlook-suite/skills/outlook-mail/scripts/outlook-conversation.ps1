param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [int]$Limit = 100,
    [ValidateSet("domain", "full", "none")]
    [string]$Links = "domain",
    [switch]$TrimQuoted,
    [int]$Preview = 0,
    [switch]$HeadersOnly,
    [switch]$Json
)

# Outlook Conversation: the whole thread an email belongs to, oldest first, across folders
# (Inbox, Sent Items, ...). Every message is shown in full: headers, EntryID, recipients,
# body and links (as [URL-001] references with a table per message).
#
# Token savers (optional):
#   -TrimQuoted   cut each body at the quoted history ("From: ... Sent: ..."), which repeats
#                 earlier messages already shown above it
#   -Preview 300  show only the first N characters of each body
#   -HeadersOnly  no bodies
#
# Usage: .\outlook-conversation.ps1 -EntryID "00000000..."
# Usage: .\outlook-conversation.ps1 -EntryID "00000000..." -TrimQuoted

. (Join-Path $PSScriptRoot "_common.ps1")

# Cut a plain-text body where the quoted reply history starts.
function Remove-QuotedHistory([string]$text) {
    if (-not $text) { return $text }
    $patterns = @(
        '(?m)^_{10,}\s*$',                                        # Outlook separator line
        '(?m)^\s*From:\s.+\r?\n\s*Sent:\s',                       # Outlook header block
        '(?m)^\s*From:\s.+\r?\n\s*Date:\s',
        '(?m)^\s*-{2,}\s*Original Message\s*-{2,}',
        '(?m)^On .{5,200}wrote:\s*$'                              # Gmail/Apple style
    )
    $cut = $text.Length
    foreach ($p in $patterns) {
        $m = [regex]::Match($text, $p)
        if ($m.Success -and $m.Index -gt 0 -and $m.Index -lt $cut) { $cut = $m.Index }
    }
    if ($cut -lt $text.Length) { return $text.Substring(0, $cut).TrimEnd() + "`n[quoted history trimmed]" }
    return $text
}

try {
    $null = Connect-Outlook
    $target = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days

    $ids = New-Object System.Collections.Generic.List[object]
    $source = "conversation"
    try {
        $conv = $target.GetConversation()
        if ($conv) {
            $table = $conv.GetTable()
            $table.Columns.Add($script:Prop.LongTermEntryID) | Out-Null
            while (-not $table.EndOfTable) {
                $row = $table.GetNextRow()
                $eid = $row.Item($script:Prop.LongTermEntryID)
                $eid = if ($eid -is [byte[]]) { ConvertTo-HexId $eid } else { [string]$row.Item("EntryID") }
                $ids.Add($eid)
            }
        }
    } catch { }

    if ($ids.Count -eq 0) {
        # Fallback for stores without conversation support: same base subject in this folder.
        $source = "subject match (conversation view not available for this store)"
        $base = $target.Subject -replace '^((RE|FW|FWD|AW|WG)\s*:\s*)+', ''
        $folder = $target.Parent
        $since = Format-JetDate (Get-Date).AddDays(-$Days)
        $items = $folder.Items.Restrict("[ReceivedTime] >= '$since'")
        foreach ($it in $items) {
            if (($it.Subject -replace '^((RE|FW|FWD|AW|WG)\s*:\s*)+', '') -eq $base) { $ids.Add($it.EntryID) }
        }
        if ($ids -notcontains $target.EntryID) { $ids.Add($target.EntryID) }
    }

    $records = New-Object System.Collections.Generic.List[object]
    foreach ($id in $ids) {
        try {
            $it = $Namespace.GetItemFromID($id)
            if ($HeadersOnly) { $rec = ConvertTo-MailRecord $it -WithRecipients }
            else {
                $rec = ConvertTo-MailRecord $it -WithRecipients -WithBody -Links $Links
                if ($TrimQuoted) { $rec.Body = Remove-QuotedHistory $rec.Body }
                if ($Preview -gt 0 -and $rec.Body.Length -gt $Preview) { $rec.Body = $rec.Body.Substring(0, $Preview) + "..." }
            }
            $records.Add($rec)
        } catch { }
    }
    $sorted = @($records | Sort-Object { if ($_.Date) { [datetime]$_.Date } else { [datetime]::MaxValue } })
    $truncated = $sorted.Count -gt $Limit
    if ($truncated) { $sorted = @($sorted | Select-Object -Last $Limit) }

    if ($Json) {
        Write-Json ([ordered]@{
            Topic = $target.ConversationTopic; ConversationID = $target.ConversationID; Source = $source
            Count = $sorted.Count; Truncated = $truncated; Messages = @($sorted | ForEach-Object { ConvertTo-JsonRecord $_ })
        })
        exit 0
    }

    Write-Host "=== CONVERSATION ===" -ForegroundColor Cyan
    Write-Host "Topic: $($target.ConversationTopic)" -ForegroundColor Yellow
    Write-Host "ConversationID: $($target.ConversationID)" -ForegroundColor DarkGray
    Write-Host "Messages: $($sorted.Count)$(if ($truncated) { " (newest $Limit; raise -Limit for all)" })   Source: $source" -ForegroundColor Gray
    $n = 0
    foreach ($r in $sorted) {
        $n++
        $mark = if ($r.EntryID -eq $target.EntryID) { "  <- requested email" } else { "" }
        Write-Host ""
        Write-MailRecord $r -Header "--- Message $n of $($sorted.Count)$mark ---" -LinkTable
    }
    Write-Host ""
    Write-Host "=== END OF CONVERSATION ===" -ForegroundColor Cyan
}
catch {
    Stop-WithError $_
}
