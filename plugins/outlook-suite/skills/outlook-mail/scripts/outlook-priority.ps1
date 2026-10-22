param(
    [int]$Days = 1,                  # -Days 0 = no date limit
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [switch]$All,                    # no date limit (an explicit -Days still wins)
    [switch]$Unread,
    [switch]$Read,
    [string]$Category = "",          # "@Action,@Review": any of these tags
    [switch]$Untagged,
    [string]$Priority = "",          # "High", "High,Normal", "Low", "None" (not scored)
    [string]$From = "",
    [string]$Subject = "",
    [string]$Folder = "Inbox",
    [switch]$Why,                    # add Copilot's reason for each row shown (one read per email)
    [switch]$Summary,                # counts per priority level only
    [int]$Limit = 100,
    [switch]$Json
)

# Outlook Priority: mail ranked by Copilot's inbox priority (High > Normal > Low > not scored),
# newest first within each level. The priority is what Copilot shows as the up/down arrow in
# Outlook; it is stored on each message (CopilotInboxScore: 5 high, 3 normal, 1 low).
# Copilot's one-line headline is shown when it wrote one; -Why adds its reason.
#
# Usage: .\outlook-priority.ps1 -Days 1 -Unread
# Usage: .\outlook-priority.ps1 -Days 7 -Priority High -Why
# Usage: .\outlook-priority.ps1 -Days 30 -Summary
# Usage: .\outlook-priority.ps1 -Category "@Action" -Priority High,Normal -Json

. (Join-Path $PSScriptRoot "_common.ps1")

$LastVerb = "http://schemas.microsoft.com/mapi/proptag/0x10810003"   # 102 reply, 103 reply all, 104 forward
$Keywords = '"urn:schemas-microsoft-com:office:office#Keywords"'
$Levels = @("High", "Normal", "Low", "None")

function Split-List([string]$s) { return @($s -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

try {
    if ($Unread -and $Read) { throw "-Unread and -Read cannot be used together." }
    if ($Untagged -and $Category) { throw "-Untagged and -Category cannot be used together." }
    $wantPriority = @(Split-List $Priority)
    foreach ($p in $wantPriority) { if ($p -notin $Levels) { throw "-Priority takes High, Normal, Low or None (got '$p')." } }
    $wantTags = @(Split-List $Category)

    $since = $null; $until = $null
    if ($DateFrom -or $DateTo) {
        $range = Resolve-DateRange -Days $(if ($Days -gt 0) { $Days } else { 1 }) -DateFrom $DateFrom -DateTo $DateTo
        $since = $range[0]; $until = $range[1]
    } elseif ($Days -gt 0 -and -not ($All -and -not $PSBoundParameters.ContainsKey("Days"))) {
        $since = (Get-Date).AddDays(-$Days)
    }

    $null = Connect-Outlook
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder

    $q = @()
    $d = '"urn:schemas:httpmail:datereceived"'
    if ($since) { $q += "$d >= '$($since.ToUniversalTime().ToString('g'))'" }
    if ($until) { $q += "$d <= '$($until.ToUniversalTime().ToString('g'))'" }
    if ($Unread) { $q += '"urn:schemas:httpmail:read" = 0' }
    if ($Read) { $q += '"urn:schemas:httpmail:read" = 1' }
    if ($Untagged) { $q += "$Keywords IS NULL" }
    if ($wantTags.Count) { $q += "(" + (($wantTags | ForEach-Object { "$Keywords LIKE '%$(ConvertTo-DaslText $_)%'" }) -join " OR ") + ")" }
    if ($Subject) { $q += """urn:schemas:httpmail:subject"" LIKE '%$(ConvertTo-DaslText $Subject)%'" }
    if ($From) {
        $s = ConvertTo-DaslText $From
        $q += "(""urn:schemas:httpmail:fromname"" LIKE '%$s%' OR ""$($script:Prop.SenderSmtp)"" LIKE '%$s%')"
    }
    $table = if ($q.Count) { $target.GetTable("@SQL=" + ($q -join " AND ")) } else { $target.GetTable() }
    $table.Columns.RemoveAll()
    $cols = @($script:Prop.LongTermEntryID, "Subject", "SenderName", $script:Prop.SenderSmtp, "ReceivedTime",
        $script:Prop.MessageFlags, "Categories", $LastVerb, $script:CopilotProp.Score, $script:CopilotProp.Headline,
        $script:Prop.ConversationID)
    foreach ($c in $cols) { [void]$table.Columns.Add($c) }
    $table.Sort("ReceivedTime", $true)
    $rowCount = $table.GetRowCount()
    $rows = $null
    if ($rowCount -gt 0) { $rows = $table.GetArray($rowCount) }

    $list = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $rowCount; $i++) {
        $tags = @(Split-List ([string]$rows[$i, 6]))
        if ($wantTags.Count -and -not ($tags | Where-Object { $_ -in $wantTags })) { continue }
        $prio = ConvertTo-CopilotPriority $rows[$i, 8]
        $level = if ($prio) { $prio } else { "None" }
        if ($wantPriority.Count -and $level -notin $wantPriority) { continue }
        $verb = $rows[$i, 7]; $verb = if ($null -ne $verb) { [int]$verb } else { 0 }
        $eid = $rows[$i, 0]; $conv = $rows[$i, 10]
        $name = [string]$rows[$i, 2]; $smtp = [string]$rows[$i, 3]
        $rec = [ordered]@{
            EntryID         = $(if ($eid -is [byte[]]) { ConvertTo-HexId $eid } else { "" })
            ConversationID  = $(if ($conv -is [byte[]]) { ConvertTo-HexId $conv } else { "" })
            CopilotPriority = $prio
            Date            = [datetime]$rows[$i, 4]
            From            = $(if ($smtp -and $smtp -ne $name) { "$name <$smtp>" } else { $name })
            Subject         = [string]$rows[$i, 1]
            CopilotHeadline = [string]$rows[$i, 9]
            CopilotReason   = ""
            Categories      = $tags
            Unread          = -not ([int]$rows[$i, 5] -band 1)
            Replied         = $verb -in 102, 103
        }
        $list.Add($rec)
    }

    # Rank: High, Normal, Low, not scored; the table order (newest first) is kept inside each level.
    $rank = @{ "High" = 0; "Normal" = 1; "Low" = 2; "" = 3 }
    $records = $list.ToArray()
    $records = @(for ($r = 0; $r -le 3; $r++) { $records | Where-Object { $rank[$_.CopilotPriority] -eq $r } })
    $counts = [ordered]@{
        High   = @($records | Where-Object { $_.CopilotPriority -eq "High" }).Count
        Normal = @($records | Where-Object { $_.CopilotPriority -eq "Normal" }).Count
        Low    = @($records | Where-Object { $_.CopilotPriority -eq "Low" }).Count
        None   = @($records | Where-Object { -not $_.CopilotPriority }).Count
    }
    $shown = if ($Summary) { @() } else { @($records | Select-Object -First $Limit) }
    if ($Why) {
        foreach ($r in $shown) {
            try { $r.CopilotReason = (Get-CopilotInsight ($Namespace.GetItemFromID($r.EntryID))).Reason } catch { }
        }
    }

    $filters = @()
    $filters += if ($since) { "since $(Format-MailDate $since)" } else { "any date" }
    if ($until) { $filters += "until $(Format-MailDate $until)" }
    $filters += if ($Unread) { "unread" } elseif ($Read) { "read" } else { "read and unread" }
    if ($Untagged) { $filters += "no tag" } elseif ($wantTags.Count) { $filters += "tag: $($wantTags -join ' or ')" }
    if ($wantPriority.Count) { $filters += "priority: $($wantPriority -join ' or ')" }
    if ($From) { $filters += "from: '$From'" }
    if ($Subject) { $filters += "subject: '$Subject'" }

    if ($Json) {
        $items = @(foreach ($r in $shown) { ConvertTo-JsonRecord $r })
        Write-Json ([ordered]@{ folder = $target.FolderPath; filters = $filters; matched = $records.Count; counts = $counts; items = $items })
        exit 0
    }

    Write-Host "=== OUTLOOK PRIORITY (Copilot) ===" -ForegroundColor Cyan
    Write-Host "Folder: $($target.FolderPath)"
    Write-Host "Filters: $($filters -join ' | ')" -ForegroundColor Gray
    Write-Host ""
    $level = $null; $n = 0
    foreach ($r in $shown) {
        $n++
        $l = if ($r.CopilotPriority) { $r.CopilotPriority } else { "None" }
        if ($l -ne $level) {
            $level = $l
            $label = if ($l -eq "None") { "NOT SCORED" } else { $l.ToUpper() }
            Write-Host "### $label ($($counts[$l]))" -ForegroundColor Green
        }
        $state = @($(if ($r.Unread) { "UNREAD" } else { "read" }), $(if ($r.Replied) { "replied" } else { "not replied" }))
        Write-Host "[$n] $(Format-MailDate $r.Date)  $($state -join ' | ')" -ForegroundColor Cyan
        Write-Host "  Subject: $($r.Subject)" -ForegroundColor Yellow
        Write-Host "  From: $($r.From)"
        if ($r.CopilotHeadline) { Write-Host "  Copilot: $($r.CopilotHeadline)" -ForegroundColor Magenta }
        if ($r.CopilotReason) { Write-Host "  Why: $($r.CopilotReason)" -ForegroundColor Magenta }
        if ($r.Categories.Count) { Write-Host "  Tags: $($r.Categories -join ', ')" -ForegroundColor Green }
        Write-Host "  EntryID: $($r.EntryID)" -ForegroundColor DarkGray
        Write-Host ""
    }
    if ($records.Count -eq 0) { Write-Host "Nothing matches. Try a larger -Days, -All, or fewer filters." -ForegroundColor Gray }
    $more = if (-not $Summary -and $records.Count -gt $shown.Count) { " (raise -Limit to see all)" } else { "" }
    Write-Host ("=== {0} matched{1} | High {2} | Normal {3} | Low {4} | not scored {5} ===" -f `
            $records.Count, $more, $counts.High, $counts.Normal, $counts.Low, $counts.None) -ForegroundColor Cyan
    if (-not $Summary) { Write-Host "Next: outlook-read-body.ps1 -EntryID <id> | outlook-reply.ps1 -EntryID <id> | -Why for Copilot's reasons" -ForegroundColor DarkGray }
}
catch {
    Stop-WithError $_
}
