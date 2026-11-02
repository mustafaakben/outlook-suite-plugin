param(
    [Parameter(Position = 0)]
    [ValidateSet("untagged", "tagged", "reply", "all")]
    [string]$View = "untagged",
    [int]$Days = -1,                 # -1 = the view's default (untagged/tagged 3, reply 7)
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [switch]$All,                    # no date limit, read and unread
    [switch]$Unread,
    [switch]$Read,
    [string]$Category = "",          # "@Action,@Review": any of these tags
    [string]$ExcludeCategory = "",   # "@Done,@NoAction": drop mail with any of these tags
    [switch]$Untagged,
    [switch]$NotReplied,
    [switch]$Replied,
    [string]$Priority = "",          # Copilot priority: "High", "High,Normal", "Low"
    [string]$From = "",
    [string]$Subject = "",
    [string]$Folder = "Inbox",
    [ValidateSet("", "Newest", "Oldest", "Priority")]
    [string]$Sort = "",
    [int]$Limit = 100,
    [switch]$Copilot,                # add Copilot's reason and suggested replies (read per email shown)
    [switch]$IncludeBody,            # add each shown email's text (read per email shown)
    [int]$BodyChars = 4000,          # with -IncludeBody: cut each body after this many characters; 0 = no cap
    [switch]$Json
)

# Outlook Triage: facts for deciding what to do with mail, no rules and no AI.
#   untagged  (default) unread mail with no category, last 3 days: what nobody has sorted yet
#   tagged    mail with a category, last 3 days, read and unread, grouped by its tags
#   reply     mail you have not replied to, last 7 days, any tag or none, oldest first
#   all       every email of the last 3 days: read or unread, tagged or not, replied or not
#             (narrow it with any filter, e.g. -ExcludeCategory "@Done")
# To process each email exactly once, do not use unread as "not processed yet": opening an email
# anywhere marks it read, so it would be skipped. Tag the mail you have handled and ask for
# all -ExcludeCategory "<handled tags>", with -Days covering the time since the last run.
# Every filter works with every view and overrides the view's default.
# Each row: date, sender, exact subject, tags, unread, replied, Copilot priority and
# headline, whether you are in To or only CC, recipient count, flag due date, EntryID,
# Message-ID and WebLink (opens the email in Outlook on the web; empty for non-Exchange mail).
# -ExcludeCategory drops mail that has any of the given tags (the untagged view has none anyway).
# -IncludeBody adds each shown email's text, links as [URL-001], cut at -BodyChars (default 4000).
# -Copilot adds Copilot's reason for the priority and its suggested replies (slower: per email).
#
# Usage: .\outlook-triage.ps1
# Usage: .\outlook-triage.ps1 untagged -Days 7 -All
# Usage: .\outlook-triage.ps1 tagged -Category "@Action,@Review" -Unread
# Usage: .\outlook-triage.ps1 reply -Category "!Urgent,@Action" -Priority High -Json
# Usage: .\outlook-triage.ps1 tagged -All -Days 7 -ExcludeCategory "@Done,@NoAction"
# Usage: .\outlook-triage.ps1 all -Days 3 -ExcludeCategory "@Done,@NoAction" -IncludeBody -Json
# Usage: .\outlook-triage.ps1 untagged -IncludeBody -BodyChars 2000 -Limit 20 -Json
# Usage: .\outlook-triage.ps1 tagged -All -Folder "Inbox/Projects"
# Usage: .\outlook-triage.ps1 reply -Priority High -Copilot

. (Join-Path $PSScriptRoot "_common.ps1")

$LastVerb = "http://schemas.microsoft.com/mapi/proptag/0x10810003"   # 102 reply, 103 reply all, 104 forward
$FlagDue = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040"
$ToMe = "http://schemas.microsoft.com/mapi/proptag/0x0057000B"      # Outlook's own "sent to me" / "CC to me" flags
$CcMe = "http://schemas.microsoft.com/mapi/proptag/0x0058000B"
$Keywords = '"urn:schemas-microsoft-com:office:office#Keywords"'
$VerbNames = @{ 102 = "replied"; 103 = "replied all"; 104 = "forwarded" }

function Split-List([string]$s) { return @($s -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

# Body text for a batch: links as [URL-00N] (the same numbers outlook-read-body and outlook-links
# use), invisible padding and runs of blank lines removed, cut at $Max characters (0 = no cap).
function Format-TriageBody([string]$Text, [int]$Max) {
    $t = (Format-MailLinks -Text $Text -Mode domain).Text
    $t = ($t -replace $script:InvisibleChars, '' -replace '\r\n?', "`n" -replace '[^\S\n]+\n', "`n" -replace '\n{3,}', "`n`n").Trim()
    if ($Max -gt 0 -and $t.Length -gt $Max) {
        return @{ Text = $t.Substring(0, $Max) + "`n[... cut at $Max of $($t.Length) characters; full text: outlook-read-body.ps1 -EntryID <EntryID>]"; Truncated = $true }
    }
    return @{ Text = $t; Truncated = $false }
}

try {
    if ($Unread -and $Read) { throw "-Unread and -Read cannot be used together." }
    if ($Replied -and $NotReplied) { throw "-Replied and -NotReplied cannot be used together." }
    if ($Untagged -and $Category) { throw "-Untagged and -Category cannot be used together." }
    $wantPriority = @(Split-List $Priority)
    foreach ($p in $wantPriority) { if ($p -notin "High", "Normal", "Low") { throw "-Priority takes High, Normal or Low (got '$p')." } }
    $wantTags = @(Split-List $Category)
    $skipTags = @(Split-List $ExcludeCategory)
    $both = @($wantTags | Where-Object { $_ -in $skipTags })
    if ($both.Count -gt 0) { throw "-Category and -ExcludeCategory both name '$($both -join "', '")'." }
    if ($BodyChars -lt 0) { throw "-BodyChars takes 0 (no cap) or a positive number of characters." }

    # View defaults; any filter given on the command line wins.
    $readState = ""; $tagState = ""; $replyState = ""; $sortBy = "Newest"
    switch ($View) {
        "untagged" { $defaultDays = 3; $readState = "unread"; $tagState = "untagged" }
        "tagged" { $defaultDays = 3; $tagState = "tagged" }
        "reply" { $defaultDays = 7; $replyState = "notreplied"; $sortBy = "Oldest" }
        "all" { $defaultDays = 3 }
    }
    if ($All) { $readState = "" }
    if ($Unread) { $readState = "unread" }
    if ($Read) { $readState = "read" }
    if ($Untagged) { $tagState = "untagged" }
    if ($wantTags.Count -gt 0) { $tagState = "tags" }
    if ($NotReplied) { $replyState = "notreplied" }
    if ($Replied) { $replyState = "replied" }
    if ($Sort) { $sortBy = $Sort }

    $since = $null; $until = $null
    if ($DateFrom -or $DateTo) {
        $range = Resolve-DateRange -Days $(if ($Days -ge 0) { $Days } else { $defaultDays }) -DateFrom $DateFrom -DateTo $DateTo
        $since = $range[0]; $until = $range[1]
    } elseif ($Days -gt 0 -or (-not $All -and $Days -ne 0)) {
        # An explicit -Days wins over -All; -All alone (or -Days 0) means no date limit.
        $since = (Get-Date).AddDays(-$(if ($Days -gt 0) { $Days } else { $defaultDays }))
    }

    $null = Connect-Outlook
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
    $me = [string]$Namespace.CurrentUser.Name

    # Everything Outlook can filter on goes into one indexed table query.
    $q = @()
    $d = '"urn:schemas:httpmail:datereceived"'
    if ($since) { $q += "$d >= '$($since.ToUniversalTime().ToString('g'))'" }
    if ($until) { $q += "$d <= '$($until.ToUniversalTime().ToString('g'))'" }
    if ($readState -eq "unread") { $q += '"urn:schemas:httpmail:read" = 0' }
    if ($readState -eq "read") { $q += '"urn:schemas:httpmail:read" = 1' }
    if ($tagState -eq "untagged") { $q += "$Keywords IS NULL" }
    if ($tagState -eq "tagged") { $q += "NOT($Keywords IS NULL)" }
    if ($tagState -eq "tags") { $q += "(" + (($wantTags | ForEach-Object { "$Keywords LIKE '%$(ConvertTo-DaslText $_)%'" }) -join " OR ") + ")" }
    if ($Subject) { $q += """urn:schemas:httpmail:subject"" LIKE '%$(ConvertTo-DaslText $Subject)%'" }
    if ($From) {
        $s = ConvertTo-DaslText $From
        $q += "(""urn:schemas:httpmail:fromname"" LIKE '%$s%' OR ""$($script:Prop.SenderSmtp)"" LIKE '%$s%')"
    }
    $table = if ($q.Count) { $target.GetTable("@SQL=" + ($q -join " AND ")) } else { $target.GetTable() }
    $table.Columns.RemoveAll()
    $cols = @($script:Prop.LongTermEntryID, "Subject", "SenderName", $script:Prop.SenderSmtp, "ReceivedTime",
        $script:Prop.MessageFlags, "urn:schemas:httpmail:displayto", "urn:schemas:httpmail:displaycc", "Categories",
        $script:Prop.FlagStatus, $FlagDue, $LastVerb, $script:CopilotProp.Score, $script:Prop.ConversationID, $script:Prop.Preview,
        $ToMe, $CcMe, $script:CopilotProp.Headline, $script:Prop.MessageID)
    foreach ($c in $cols) { [void]$table.Columns.Add($c) }
    $table.Sort("ReceivedTime", $true)
    $rowCount = $table.GetRowCount()
    $rows = $null
    if ($rowCount -gt 0) { $rows = $table.GetArray($rowCount) }

    # The rest (exact tags, replied state, Copilot priority) is checked on the rows.
    $now = Get-Date
    $list = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $rowCount; $i++) {
        $tags = @(Split-List ([string]$rows[$i, 8]))
        if ($tagState -eq "tags" -and -not ($tags | Where-Object { $_ -in $wantTags })) { continue }
        # Exclusions are checked here, not in the query: a DASL NOT LIKE would also drop untagged mail.
        if ($skipTags.Count -gt 0 -and ($tags | Where-Object { $_ -in $skipTags })) { continue }
        $verb = $rows[$i, 11]; $verb = if ($null -ne $verb) { [int]$verb } else { 0 }
        $isReplied = $verb -in 102, 103
        if ($replyState -eq "notreplied" -and $isReplied) { continue }
        if ($replyState -eq "replied" -and -not $isReplied) { continue }
        $prio = ConvertTo-CopilotPriority $rows[$i, 12]
        if ($wantPriority.Count -gt 0 -and $prio -notin $wantPriority) { continue }

        $toNames = @(Split-List ([string]$rows[$i, 6])); $ccNames = @(Split-List ([string]$rows[$i, 7]))
        $you = if ($rows[$i, 15]) { "To" } elseif ($rows[$i, 16]) { "CC" }
        elseif ($null -eq $rows[$i, 15]) {
            # Flags missing (some item types): fall back to the display names.
            if ($toNames -contains $me) { "To" } elseif ($ccNames -contains $me) { "CC" } else { "" }
        } else { "" }
        $flagStatus = [int]$rows[$i, 9]
        $due = $rows[$i, 10]; $dueText = if ($flagStatus -eq 2 -and $due -and ([datetime]$due).Year -lt 4500) { ([datetime]$due).ToString("yyyy-MM-dd") } else { "" }
        $eid = $rows[$i, 0]; $conv = $rows[$i, 13]
        $name = [string]$rows[$i, 2]; $smtp = [string]$rows[$i, 3]
        $date = [datetime]$rows[$i, 4]
        $rec = [ordered]@{
            EntryID         = $(if ($eid -is [byte[]]) { ConvertTo-HexId $eid } else { "" })
            ConversationID  = $(if ($conv -is [byte[]]) { ConvertTo-HexId $conv } else { "" })
            MessageID       = [string]$rows[$i, 18]
            WebLink         = ""                      # built below, only for the rows shown
            Date            = $date
            From            = $(if ($smtp -and $smtp -ne $name) { "$name <$smtp>" } else { $name })
            Subject         = [string]$rows[$i, 1]
            Categories      = $tags
            Unread          = -not ([int]$rows[$i, 5] -band 1)
            Replied         = $isReplied
            LastAction      = $(if ($VerbNames.ContainsKey($verb)) { $VerbNames[$verb] } else { "" })
            CopilotPriority = $prio
            CopilotHeadline = [string]$rows[$i, 17]
            You             = $you
            Recipients      = $toNames.Count + $ccNames.Count
            Flag            = $(if ($flagStatus -eq 2) { "Flagged" } elseif ($flagStatus -eq 1) { "Complete" } else { "" })
            FlagDue         = $dueText
            DaysWaiting     = [int][Math]::Floor(($now - $date).TotalDays)
            Preview         = [string]$rows[$i, 14]   # formatted below, only for the rows shown
        }
        $list.Add($rec)
    }

    $records = $list.ToArray()
    if ($sortBy -eq "Oldest") { [array]::Reverse($records) }
    if ($sortBy -eq "Priority") {
        $rank = @{ "High" = 0; "Normal" = 1; "Low" = 2; "" = 3 }
        $records = @($records | Sort-Object @{ Expression = { $rank[$_.CopilotPriority] } }, @{ Expression = { $_.Date }; Descending = $true })
    }
    $matched = $records.Count
    if ($View -eq "tagged") {
        # Group by the tag set as Outlook shows it; the groups keep the sort order inside.
        $records = @($records | Group-Object { $_.Categories -join ", " } | Sort-Object Name | ForEach-Object { $_.Group })
    }
    $shown = @($records | Select-Object -First $Limit)
    foreach ($r in $shown) {
        $r.Preview = Format-Preview $r.Preview 160
        $r.WebLink = Get-MailWebLink -Namespace $Namespace -EntryID $r.EntryID -StoreID $target.StoreID
        # Table text columns stop at 255 characters (long Message-IDs, e.g. Teams notices): read the item.
        if ($r.MessageID.Length -ge 255) {
            try { $r.MessageID = [string]$Namespace.GetItemFromID($r.EntryID, $target.StoreID).PropertyAccessor.GetProperty($script:Prop.MessageID) } catch { }
        }
    }
    if ($Copilot -or $IncludeBody) {
        # Copilot's reason, suggested replies and the body are not table columns: one read per email shown.
        foreach ($r in $shown) {
            if ($Copilot) { $r.CopilotReason = ""; $r.SmartReplies = @() }
            if ($IncludeBody) { $r.Body = ""; $r.BodyTruncated = $false }
            $item = $null
            try { $item = $Namespace.GetItemFromID($r.EntryID, $target.StoreID) } catch { continue }
            if ($Copilot) {
                try { $ci = Get-CopilotInsight $item; $r.CopilotReason = $ci.Reason; $r.SmartReplies = @($ci.SmartReplies) } catch { }
            }
            if ($IncludeBody) {
                try { $b = Format-TriageBody ([string]$item.Body) $BodyChars; $r.Body = $b.Text; $r.BodyTruncated = $b.Truncated } catch { }
            }
        }
    }

    $filters = @()
    $filters += if ($since) { "since $(Format-MailDate $since)" } else { "any date" }
    if ($until) { $filters += "until $(Format-MailDate $until)" }
    $filters += switch ($readState) { "unread" { "unread" } "read" { "read" } default { "read and unread" } }
    $filters += switch ($tagState) { "untagged" { "no tag" } "tagged" { "any tag" } "tags" { "tag: $($wantTags -join ' or ')" } default { "tagged or not" } }
    if ($skipTags.Count -gt 0) { $filters += "without tag: $($skipTags -join ', ')" }
    if ($replyState) { $filters += $(if ($replyState -eq "replied") { "replied" } else { "not replied" }) }
    if ($wantPriority.Count) { $filters += "Copilot priority: $($wantPriority -join ' or ')" }
    if ($From) { $filters += "from: '$From'" }
    if ($Subject) { $filters += "subject: '$Subject'" }

    $summary = [ordered]@{
        matched    = $matched
        shown      = $shown.Count
        unread     = @($records | Where-Object { $_.Unread }).Count
        untagged   = @($records | Where-Object { $_.Categories.Count -eq 0 }).Count
        notReplied = @($records | Where-Object { -not $_.Replied }).Count
        copilotHigh = @($records | Where-Object { $_.CopilotPriority -eq "High" }).Count
        flagged    = @($records | Where-Object { $_.Flag -eq "Flagged" }).Count
    }

    if ($Json) {
        $items = @(foreach ($r in $shown) { ConvertTo-JsonRecord $r })
        Write-Json ([ordered]@{ view = $View; folder = $target.FolderPath; filters = $filters; sort = $sortBy; summary = $summary; items = $items })
        exit 0
    }

    Write-Host "=== OUTLOOK TRIAGE: $View ===" -ForegroundColor Cyan
    Write-Host "Folder: $($target.FolderPath)"
    Write-Host "Filters: $($filters -join ' | ')" -ForegroundColor Gray
    Write-Host ""
    $group = $null; $n = 0
    foreach ($r in $shown) {
        $n++
        if ($View -eq "tagged") {
            $g = $r.Categories -join ", "
            if ($g -ne $group) { $group = $g; Write-Host "### $g ($(@($records | Where-Object { ($_.Categories -join ', ') -eq $g }).Count))" -ForegroundColor Green }
        }
        $state = @($(if ($r.Unread) { "UNREAD" } else { "read" }))
        $state += if ($r.Replied) { $r.LastAction } elseif ($r.LastAction) { "$($r.LastAction), not replied" } else { "not replied" }
        if ($r.CopilotPriority) { $state += "Copilot: $($r.CopilotPriority)" }
        if (-not $r.Replied -and $View -eq "reply") { $state += "waiting $($r.DaysWaiting)d" }
        Write-Host "[$n] $(Format-MailDate $r.Date)  $($state -join ' | ')" -ForegroundColor Cyan
        Write-Host "  Subject: $($r.Subject)" -ForegroundColor Yellow
        $you = if ($r.You) { "you in $($r.You)" } else { "you not in To/CC" }
        Write-Host "  From: $($r.From)   ($you, $($r.Recipients) recipient$(if ($r.Recipients -ne 1) { 's' }))"
        if ($r.Categories.Count) { Write-Host "  Tags: $($r.Categories -join ', ')" -ForegroundColor Green }
        if ($r.Flag) { Write-Host "  Flag: $($r.Flag)$(if ($r.FlagDue) { ", due $($r.FlagDue)" })" -ForegroundColor Red }
        if ($r.CopilotHeadline) { Write-Host "  Copilot: $($r.CopilotHeadline)" -ForegroundColor Magenta }
        if ($r.Contains("CopilotReason") -and $r.CopilotReason) { Write-Host "  Copilot reason: $($r.CopilotReason)" -ForegroundColor Magenta }
        if ($r.Contains("SmartReplies") -and $r.SmartReplies.Count) {
            $k = 0; Write-Host ("  Suggested replies: " + (($r.SmartReplies | ForEach-Object { $k++; "$k) $_" }) -join "  ")) -ForegroundColor Magenta
        }
        if ($r.Contains("Body")) {
            Write-Host "  Body:" -ForegroundColor Gray
            Write-Host $r.Body
            Write-Host "  --- end of body ---" -ForegroundColor DarkGray
        } elseif ($r.Preview) { Write-Host "  Preview: $($r.Preview)" -ForegroundColor Gray }
        Write-Host "  EntryID: $($r.EntryID)" -ForegroundColor DarkGray
        if ($r.MessageID) { Write-Host "  Message-ID: $($r.MessageID)" -ForegroundColor DarkGray }
        if ($r.WebLink) { Write-Host "  WebLink: $($r.WebLink)" -ForegroundColor DarkGray }
        Write-Host ""
    }
    if ($n -eq 0) { Write-Host "Nothing matches. Try a larger -Days, -All, or fewer filters." -ForegroundColor Gray }
    $more = if ($matched -gt $shown.Count) { " (raise -Limit to see all $matched)" } else { "" }
    Write-Host ("=== {0} shown of {1}{2} | unread {3} | untagged {4} | not replied {5} | Copilot high {6} | flagged {7} ===" -f `
            $shown.Count, $matched, $more, $summary.unread, $summary.untagged, $summary.notReplied, $summary.copilotHigh, $summary.flagged) -ForegroundColor Cyan
    Write-Host "Next: outlook-read-body.ps1 -EntryID <id> | outlook-assign-category.ps1 | outlook-reply.ps1 | outlook-bulk.ps1" -ForegroundColor DarkGray
}
catch {
    Stop-WithError $_
}
