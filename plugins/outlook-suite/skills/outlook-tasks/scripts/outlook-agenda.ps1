param(
    [string]$Day = "today",
    [switch]$Week,
    [int]$MailDays = 3,
    [int]$Limit = 10,
    [switch]$Json
)

# Outlook Agenda: a read-only "what's on my plate" brief for a day (or 7 days with -Week).
# Sections, each with a count:
#   Calendar      per day, classified by the calendar skill (_calendar.ps1): meetings (others invited),
#                 holds (your own movable blocks), firm events (classes, office hours, travel, Out of
#                 Office, private) and all-day events, with the day's tier, conflicts and the 2-meeting cap
#   Tasks         due in the window, and overdue
#   Flagged email follow-ups due in the window, and overdue
#   Priority mail unread Inbox mail from the last -MailDays days that Copilot rates High or that is marked High importance
# Nothing is changed. Lists show at most -Limit items each; counts are always complete.
# Usage: .\outlook-agenda.ps1                      (today)
#        .\outlook-agenda.ps1 -Day tomorrow
#        .\outlook-agenda.ps1 -Day 2026-11-03
#        .\outlook-agenda.ps1 -Week                (today and the next 6 days)
#        .\outlook-agenda.ps1 -Day monday -Week -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")
# One meeting/hold classifier for the whole suite: the calendar skill's (Get-CalendarEvents, Get-CalendarDay).
. (Join-Path $PSScriptRoot "..\..\outlook-calendar\scripts\_calendar.ps1")

# "today", "tomorrow", a weekday name (the next one, today included) or any date.
function Resolve-AgendaDay([string]$Text) {
    $today = (Get-Date).Date
    switch -regex ($Text.Trim().ToLower()) {
        '^(today|)$' { return $today }
        '^tomorrow$' { return $today.AddDays(1) }
        '^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)$' {
            $target = [int][System.DayOfWeek]$Text.Trim()
            return $today.AddDays((($target - [int]$today.DayOfWeek) + 7) % 7)
        }
    }
    try { return ([datetime]$Text).Date } catch { throw "Cannot read -Day '$Text'. Use today, tomorrow, a weekday name or a date like 2026-11-03." }
}

# Unread Inbox mail from the last $Days days that Copilot rates High (score 5) or that is marked High importance.
function Get-PriorityUnread($Namespace, [int]$Days) {
    $inbox = $Namespace.GetDefaultFolder(6)
    $since = (Get-Date).AddDays(-$Days).ToUniversalTime().ToString("g")
    $table = $inbox.GetTable("@SQL=""urn:schemas:httpmail:read"" = 0 AND ""urn:schemas:httpmail:datereceived"" >= '$since'")
    $table.Columns.RemoveAll()
    foreach ($c in @("EntryID", "Subject", "SenderName", "ReceivedTime", "Importance", $script:CopilotProp.Score, $script:CopilotProp.Headline, $script:Prop.LongTermEntryID)) { [void]$table.Columns.Add($c) }
    $table.Sort("ReceivedTime", $true)
    $n = $table.GetRowCount()
    $rows = $null
    if ($n -gt 0) { $rows = $table.GetArray($n) }   # direct assignment keeps the 2-D array
    $list = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $n; $i++) {
        $copilot = ConvertTo-CopilotPriority $rows[$i, 5]
        $high = [int]$rows[$i, 4] -eq 2
        if ($copilot -ne "High" -and -not $high) { continue }
        $eid = $rows[$i, 7]; $eid = if ($eid -is [byte[]]) { ConvertTo-HexId $eid } else { [string]$rows[$i, 0] }
        $why = @(); if ($copilot -eq "High") { $why += "Copilot High" }; if ($high) { $why += "High importance" }
        $rec = [ordered]@{
            EntryID  = $eid
            Subject  = [string]$rows[$i, 1]
            From     = [string]$rows[$i, 2]
            Date     = $rows[$i, 3]
            Why      = ($why -join ", ")
            Headline = [string]$rows[$i, 6]
        }
        $list.Add($rec)
    }
    return , $list
}

function Format-EventTime($e) {
    if ($e.AllDay) { return "all day" }
    return "$($e.Start.ToString('HH:mm')) to $($e.End.ToString('HH:mm'))"
}

try {
    $first = Resolve-AgendaDay $Day
    $nDays = if ($Week) { 7 } else { 1 }
    $from = $first; $to = $first.AddDays($nDays)
    $today = (Get-Date).Date

    $null = Connect-Outlook
    $events = Get-CalendarEvents -Namespace $Namespace -From $from -To $to
    Set-CalendarConflicts $events
    $todo = (Get-TaskRecords -Namespace $Namespace -Flagged).Records   # open tasks + flagged mail, by due date
    $priority = Get-PriorityUnread $Namespace $MailDays

    $inWindow = { param($r) $r.Due -and $r.Due.Date -ge $from -and $r.Due.Date -lt $to }
    $isOverdue = { param($r) $r.Due -and $r.Due.Date -lt $today }
    $tasksDue = @($todo | Where-Object { $_.Type -eq "Task" -and (& $inWindow $_) })
    $tasksOver = @($todo | Where-Object { $_.Type -eq "Task" -and (& $isOverdue $_) })
    $mailDue = @($todo | Where-Object { $_.Type -eq "Mail" -and (& $inWindow $_) })
    $mailOver = @($todo | Where-Object { $_.Type -eq "Mail" -and (& $isOverdue $_) })

    $days = @(for ($d = 0; $d -lt $nDays; $d++) {
        $dayStart = $first.AddDays($d); $dayEnd = $dayStart.AddDays(1)
        $info = Get-CalendarDay -Records $events -Date $dayStart   # excludes cancelled and declined
        $live = @($info.Events)
        $flags = @()
        if ($info.OverCap) { $flags += "over the 2-meeting cap ($($info.Meetings) meetings)" }
        if ($info.Conflicts -gt 0) { $flags += "$($info.Conflicts) event(s) in a conflict" }
        if ($info.Away) { $flags += "away (all-day Out of Office or Busy)" }
        [ordered]@{
            Date      = $dayStart
            Tier      = $info.TierName
            Counts    = [ordered]@{ Meetings = $info.Meetings; Holds = $info.Holds; Firm = $info.Firm; AllDay = @($info.AllDay).Count }
            Meetings  = @($live | Where-Object { $_.Kind -eq "meeting" })
            Firm      = @($live | Where-Object { $_.Kind -eq "firm" })
            Holds     = @($live | Where-Object { $_.Kind -eq "hold" })
            AllDay    = @($live | Where-Object { $_.Kind -eq "allday" })
            Cancelled = @($events | Where-Object { $_.Start -lt $dayEnd -and $_.End -gt $dayStart -and ($_.Cancelled -or $_.Response -eq "Declined") })
            Flags     = $flags
        }
    })
    $counts = [ordered]@{
        Meetings = [int]($days | ForEach-Object { $_.Counts.Meetings } | Measure-Object -Sum).Sum
        Holds = [int]($days | ForEach-Object { $_.Counts.Holds } | Measure-Object -Sum).Sum
        Firm = [int]($days | ForEach-Object { $_.Counts.Firm } | Measure-Object -Sum).Sum
        AllDay = [int]($days | ForEach-Object { $_.Counts.AllDay } | Measure-Object -Sum).Sum
        TasksDue = $tasksDue.Count; TasksOverdue = $tasksOver.Count
        FlaggedMailDue = $mailDue.Count; FlaggedMailOverdue = $mailOver.Count
        PriorityUnread = $priority.Count
    }

    if ($Json) {
        $jsonDays = @(foreach ($d in $days) {
            [ordered]@{
                Date      = $d.Date.ToString("yyyy-MM-dd")
                Tier      = $d.Tier
                Counts    = $d.Counts
                Meetings  = @($d.Meetings | ForEach-Object { ConvertTo-EventJson $_ })
                Firm      = @($d.Firm | ForEach-Object { ConvertTo-EventJson $_ })
                Holds     = @($d.Holds | ForEach-Object { ConvertTo-EventJson $_ })
                AllDay    = @($d.AllDay | ForEach-Object { ConvertTo-EventJson $_ })
                Cancelled = @($d.Cancelled | ForEach-Object { ConvertTo-EventJson $_ })
                Flags     = @($d.Flags)
            }
        })
        $taskJson = { param($list) @($list | Select-Object -First $Limit | ForEach-Object { ConvertTo-TaskJson $_ }) }
        Write-Json ([ordered]@{
            from = $from.ToString("yyyy-MM-dd"); to = $to.AddDays(-1).ToString("yyyy-MM-dd"); counts = $counts
            days = $jsonDays
            tasks = [ordered]@{ due = (& $taskJson $tasksDue); overdue = (& $taskJson $tasksOver) }
            flaggedMail = [ordered]@{ due = (& $taskJson $mailDue); overdue = (& $taskJson $mailOver) }
            priorityMail = @($priority | Select-Object -First $Limit | ForEach-Object { $o = ConvertTo-JsonRecord $_; $o })
        }) -Depth 10
        exit 0
    }

    $title = if ($Week) { "$($from.ToString('ddd, MMM d')) to $($to.AddDays(-1).ToString('ddd, MMM d, yyyy'))" } else { $from.ToString("dddd, MMM d, yyyy") }
    Write-Host "=== AGENDA: $title ===" -ForegroundColor Cyan
    Write-Host ("Meetings {0} | Firm {1} | Holds {2} | All-day {3} | Tasks due {4} (overdue {5}) | Flagged email due {6} (overdue {7}) | Priority unread {8}" -f `
        $counts.Meetings, $counts.Firm, $counts.Holds, $counts.AllDay, $counts.TasksDue, $counts.TasksOverdue, $counts.FlaggedMailDue, $counts.FlaggedMailOverdue, $counts.PriorityUnread) -ForegroundColor Gray

    Write-Host "`n--- Calendar ---" -ForegroundColor Cyan
    foreach ($d in $days) {
        $k = $d.Counts
        $sum = "$($k.Meetings) meeting(s), $($k.Firm) firm, $($k.Holds) hold(s)$(if ($k.AllDay) { ", $($k.AllDay) all-day" })"
        Write-Host "$($d.Date.ToString('ddd, MMM d')): $sum  [$($d.Tier)]" -ForegroundColor Yellow
        $labels = @{ meeting = "MEETING"; firm = "FIRM   "; hold = "hold   "; allday = "ALL-DAY" }
        foreach ($e in (@($d.AllDay) + @($d.Meetings) + @($d.Firm) + @($d.Holds) | Sort-Object { -not $_.AllDay }, { $_.Start })) {
            $extra = @()
            if ($e.Location) { $extra += $e.Location }
            if ($e.Busy -and $e.Busy -ne "Busy") { $extra += $e.Busy }
            if ($e.Kind -eq "meeting" -and $e.Organizer) { $extra += "by $($e.Organizer)" }
            if ($e.ConflictsWith -and $e.ConflictsWith.Count) { $extra += "CONFLICT" }
            Write-Host ("  {0}  {1,-14} {2}{3}" -f $labels[$e.Kind], (Format-EventTime $e), $e.Subject, $(if ($extra) { "  [" + ($extra -join "; ") + "]" } else { "" })) -ForegroundColor $(if ($e.Kind -eq "meeting") { "White" } else { "Gray" })
            if ($e.EntryID) { Write-Host "           EntryID: $($e.EntryID)" -ForegroundColor DarkGray }
        }
        foreach ($e in $d.Cancelled) { Write-Host "  cancelled/declined  $(Format-EventTime $e) $($e.Subject)" -ForegroundColor DarkGray }
        foreach ($f in $d.Flags) { Write-Host "  ! $f" -ForegroundColor Red }
    }
    $section = {
        param([string]$Name, $Due, $Over)
        Write-Host "`n--- $Name ---" -ForegroundColor Cyan
        if ($Due.Count -eq 0 -and $Over.Count -eq 0) { Write-Host "  (none)" -ForegroundColor Gray; return }
        foreach ($pair in @(@("Due in this window", $Due), @("Overdue", $Over))) {
            $list = $pair[1]
            if ($list.Count -eq 0) { continue }
            Write-Host "  $($pair[0]): $($list.Count)$(if ($list.Count -gt $Limit) { " (showing $Limit)" })" -ForegroundColor Yellow
            foreach ($r in ($list | Select-Object -First $Limit)) {
                $due = if ($r.Due) { $r.Due.ToString("ddd, MMM d") } else { "" }
                $pri = if ($r.Importance -eq "High") { " [High]" } else { "" }
                Write-Host "    $due  $($r.Subject)$pri" -ForegroundColor $(if ($pair[0] -eq "Overdue") { "Red" } else { "White" })
                Write-Host "           EntryID: $($r.EntryID)" -ForegroundColor DarkGray
            }
        }
    }
    & $section "Tasks" $tasksDue $tasksOver
    & $section "Flagged email" $mailDue $mailOver

    Write-Host "`n--- Priority unread mail (last $MailDays days) ---" -ForegroundColor Cyan
    if ($priority.Count -eq 0) { Write-Host "  (none)" -ForegroundColor Gray }
    foreach ($m in ($priority | Select-Object -First $Limit)) {
        Write-Host "  $(Format-MailDate $m.Date)  $($m.From): $($m.Subject)  [$($m.Why)]"
        if ($m.Headline) { Write-Host "           $($m.Headline)" -ForegroundColor Gray }
        Write-Host "           EntryID: $($m.EntryID)" -ForegroundColor DarkGray
    }
    if ($priority.Count -gt $Limit) { Write-Host "  ... $($priority.Count - $Limit) more (raise -Limit)" -ForegroundColor Gray }
}
catch {
    Stop-WithError $_
}
