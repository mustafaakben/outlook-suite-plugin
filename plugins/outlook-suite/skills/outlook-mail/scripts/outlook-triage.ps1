param(
    [Parameter(Position = 0)]
    [ValidateSet("untagged", "tagged", "reply")]
    [string]$View = "untagged",
    [int]$Days = -1,                 # -1 = the view's default (untagged/tagged 3, reply 7)
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [switch]$All,                    # no date limit, read and unread
    [switch]$Unread,
    [switch]$Read,
    [string]$Category = "",          # "@Action,@Review": any of these tags
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
    [switch]$Json
)

# Outlook Triage: facts for deciding what to do with mail, no rules and no AI.
#   untagged  (default) unread mail with no category, last 3 days: what nobody has sorted yet
#   tagged    mail with a category, last 3 days, read and unread, grouped by its tags
#   reply     mail you have not replied to, last 7 days, any tag or none, oldest first
# Every filter works with every view and overrides the view's default.
# Each row: date, sender, exact subject, tags, unread, replied, Copilot priority,
# whether you are in To or only CC, recipient count, flag due date, EntryID.
#
# Usage: .\outlook-triage.ps1
# Usage: .\outlook-triage.ps1 untagged -Days 7 -All
# Usage: .\outlook-triage.ps1 tagged -Category "@Action,@Review" -Unread
# Usage: .\outlook-triage.ps1 reply -Category "!Urgent,@Action" -Priority High -Json
# Usage: .\outlook-triage.ps1 tagged -All -Folder "Inbox/Projects"

. (Join-Path $PSScriptRoot "_common.ps1")

# Copilot's inbox priority (Outlook's "Marked as high/low priority by Copilot") is stored on
# each message as the named property CopilotInboxScore: 5 high, 3 normal, 1 low.
$CopilotScore = "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/CopilotInboxScore"
$LastVerb = "http://schemas.microsoft.com/mapi/proptag/0x10810003"   # 102 reply, 103 reply all, 104 forward
$FlagDue = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040"
$ToMe = "http://schemas.microsoft.com/mapi/proptag/0x0057000B"      # Outlook's own "sent to me" / "CC to me" flags
$CcMe = "http://schemas.microsoft.com/mapi/proptag/0x0058000B"
$Keywords = '"urn:schemas-microsoft-com:office:office#Keywords"'
$PriorityNames = @{ 5 = "High"; 3 = "Normal"; 1 = "Low" }
$VerbNames = @{ 102 = "replied"; 103 = "replied all"; 104 = "forwarded" }

function Split-List([string]$s) { return @($s -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

try {
    if ($Unread -and $Read) { throw "-Unread and -Read cannot be used together." }
    if ($Replied -and $NotReplied) { throw "-Replied and -NotReplied cannot be used together." }
    if ($Untagged -and $Category) { throw "-Untagged and -Category cannot be used together." }
    $wantPriority = @(Split-List $Priority)
    foreach ($p in $wantPriority) { if ($p -notin "High", "Normal", "Low") { throw "-Priority takes High, Normal or Low (got '$p')." } }
    $wantTags = @(Split-List $Category)

    # View defaults; any filter given on the command line wins.
    $readState = ""; $tagState = ""; $replyState = ""; $sortBy = "Newest"
    switch ($View) {
        "untagged" { $defaultDays = 3; $readState = "unread"; $tagState = "untagged" }
        "tagged" { $defaultDays = 3; $tagState = "tagged" }
        "reply" { $defaultDays = 7; $replyState = "notreplied"; $sortBy = "Oldest" }
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
        $script:Prop.FlagStatus, $FlagDue, $LastVerb, $CopilotScore, $script:Prop.ConversationID, $script:Prop.Preview,
        $ToMe, $CcMe)
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
        $verb = $rows[$i, 11]; $verb = if ($null -ne $verb) { [int]$verb } else { 0 }
        $isReplied = $verb -in 102, 103
        if ($replyState -eq "notreplied" -and $isReplied) { continue }
        if ($replyState -eq "replied" -and -not $isReplied) { continue }
        $score = $rows[$i, 12]; $prio = if ($null -ne $score -and $PriorityNames.ContainsKey([int]$score)) { $PriorityNames[[int]$score] } else { "" }
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
            Date            = $date
            From            = $(if ($smtp -and $smtp -ne $name) { "$name <$smtp>" } else { $name })
            Subject         = [string]$rows[$i, 1]
            Categories      = $tags
            Unread          = -not ([int]$rows[$i, 5] -band 1)
            Replied         = $isReplied
            LastAction      = $(if ($VerbNames.ContainsKey($verb)) { $VerbNames[$verb] } else { "" })
            CopilotPriority = $prio
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
    foreach ($r in $shown) { $r.Preview = Format-Preview $r.Preview 160 }

    $filters = @()
    $filters += if ($since) { "since $(Format-MailDate $since)" } else { "any date" }
    if ($until) { $filters += "until $(Format-MailDate $until)" }
    $filters += switch ($readState) { "unread" { "unread" } "read" { "read" } default { "read and unread" } }
    $filters += switch ($tagState) { "untagged" { "no tag" } "tagged" { "any tag" } "tags" { "tag: $($wantTags -join ' or ')" } default { "tagged or not" } }
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
        if ($r.Preview) { Write-Host "  Preview: $($r.Preview)" -ForegroundColor Gray }
        Write-Host "  EntryID: $($r.EntryID)" -ForegroundColor DarkGray
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
