# Helpers for the outlook-tasks scripts (tasks, and flagged mail as to-dos).
# Dot-source after _common.ps1:  . (Join-Path $PSScriptRoot "_tasks.ps1")

# Task properties as Folder.GetTable columns. The same named properties exist on tasks and
# on flagged mail, so one table query covers the Tasks folder and the To-Do List.
# Dates come back from the table in UTC (see ConvertFrom-TaskTableDate).
$script:TaskProp = @{
    Start         = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81040040"
    Due           = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040"
    Status        = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003"
    Percent       = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81020005"
    Complete      = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/811C000B"
    DateCompleted = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/810F0040"
    Recurring     = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/8126000B"
    ReminderSet   = "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/8503000B"
    ReminderTime  = "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85020040"
    LongTermID    = "http://schemas.microsoft.com/mapi/proptag/0x66700102"
    # Set by outlook-tasks-from-mail.ps1: the EntryID of the email a task was made from.
    # The /0x0000001F type suffix lets a table use the property before any item has it.
    SourceEntryID = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/OutlookSuiteSourceEntryID/0x0000001F"
    # ...and its Internet Message-ID, which stays the same when the email moves (the EntryID does not).
    SourceMessageID = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/OutlookSuiteSourceMessageID/0x0000001F"
    InternetMessageID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
}

$script:TaskStatusMap = @{ "NotStarted" = 0; "InProgress" = 1; "Completed" = 2; "Waiting" = 3; "Deferred" = 4 }
$script:TaskStatusNames = @{ 0 = "NotStarted"; 1 = "InProgress"; 2 = "Completed"; 3 = "Waiting"; 4 = "Deferred" }
$script:TaskStatusLabels = @{ 0 = "Not Started"; 1 = "In Progress"; 2 = "Completed"; 3 = "Waiting"; 4 = "Deferred" }
$script:TaskPriorityMap = @{ "Low" = 0; "Normal" = 1; "High" = 2 }

# Date-only task dates (due, start, completed) from a table, which returns them in UTC.
# Classic Outlook stores local midnight (04:00 UTC in New York): convert back to local.
# Outlook on the web, mobile, To Do and Graph store 00:00 UTC: that is the date itself, and
# converting it would give the previous evening (a task due today would show as overdue).
function ConvertFrom-TaskTableDate($Value) {
    if ($Value -isnot [datetime]) { return $null }
    if ($Value.Year -ge 4500) { return $null }
    if ($Value.TimeOfDay.Ticks -eq 0) { return [datetime]::SpecifyKind($Value.Date, [DateTimeKind]::Unspecified) }
    return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc).ToLocalTime()
}

# A real point in time (e.g. a reminder) from a table: UTC to local, no date-only rule.
function ConvertFrom-TableUtc($Value) {
    if ($Value -isnot [datetime]) { return $null }
    if ($Value.Year -ge 4500) { return $null }
    return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc).ToLocalTime()
}

# A task date (local) for display and JSON: date only when it is midnight.
function Format-TaskDate($Date) {
    if (-not $Date) { return "" }
    $d = [datetime]$Date
    if ($d.Year -ge 4500) { return "" }
    if ($d.TimeOfDay.TotalSeconds -eq 0) { return $d.ToString("yyyy-MM-dd") }
    return Format-MailDate $d
}

# Open tasks (and with -Flagged, flagged mail) as records, sorted by due date (no due date last).
# Tasks always come from the Tasks folder; -Flagged adds the flagged email from the To-Do List
# (a search folder that can take minutes to pick up new tasks, so it is used for mail only).
# One GetTable query per folder; filtering by due window, status and search happens in memory.
function Get-TaskRecords {
    param($Namespace, [switch]$Flagged, [switch]$IncludeCompleted)
    $list = New-Object System.Collections.Generic.List[object]
    $tasks = $Namespace.GetDefaultFolder(13)   # Tasks
    Read-TaskTable -Folder $tasks -IncludeCompleted:$IncludeCompleted -Want "Task" -List $list
    if ($Flagged) { Read-TaskTable -Folder $Namespace.GetDefaultFolder(28) -IncludeCompleted:$IncludeCompleted -Want "Mail" -List $list }   # To-Do List
    # Due ascending, no due date last, then most recently modified first.
    $sorted = @($list.ToArray() | Sort-Object @{ Expression = { if ($_.Due) { $_.Due } else { [datetime]::MaxValue } } }, @{ Expression = { $_.Modified }; Descending = $true })
    $label = if ($Flagged) { "$($tasks.FolderPath) + flagged email (To-Do List)" } else { $tasks.FolderPath }
    return [pscustomobject]@{ Folder = [pscustomobject]@{ FolderPath = $label }; Records = $sorted }
}

# Adds the rows of one folder to $List: -Want Task keeps tasks, -Want Mail keeps everything else (flagged items).
function Read-TaskTable {
    param($Folder, [switch]$IncludeCompleted, [string]$Want, $List)
    $filter = if ($IncludeCompleted) { "" } else { "@SQL=`"$($script:TaskProp.Complete)`" = 0" }
    $table = $null
    if ($filter) { $table = $Folder.GetTable($filter) } else { $table = $Folder.GetTable() }
    $table.Columns.RemoveAll()
    $cols = @("EntryID", "Subject", "MessageClass", $script:TaskProp.Due, $script:TaskProp.Start, $script:TaskProp.Status,
        $script:TaskProp.Percent, $script:TaskProp.Complete, "Importance", "Categories", "LastModificationTime",
        $script:TaskProp.DateCompleted, $script:TaskProp.LongTermID, $script:TaskProp.SourceEntryID,
        $script:TaskProp.Recurring, $script:TaskProp.ReminderSet, $script:TaskProp.ReminderTime, $script:TaskProp.SourceMessageID)
    foreach ($c in $cols) { [void]$table.Columns.Add($c) }
    $total = $table.GetRowCount()
    $rows = $null
    if ($total -gt 0) { $rows = $table.GetArray($total) }   # direct assignment keeps the 2-D array
    for ($i = 0; $i -lt $total; $i++) {
        $isTask = ([string]$rows[$i, 2]) -like "IPM.Task*"
        if (($Want -eq "Task") -ne $isTask) { continue }   # task requests in Tasks; tasks in the To-Do List
        $eid = $rows[$i, 12]; $eid = if ($eid -is [byte[]]) { ConvertTo-HexId $eid } else { [string]$rows[$i, 0] }
        $status = $rows[$i, 5]
        $complete = [bool]$rows[$i, 7]
        $statusName = if ($null -ne $status) { $script:TaskStatusNames[[int]$status] } else { "NotStarted" }
        if ($complete) { $statusName = "Completed" }
        $rec = [ordered]@{
            EntryID         = $eid
            Type            = $Want
            Subject         = [string]$rows[$i, 1]
            Status          = $statusName
            Due             = (ConvertFrom-TaskTableDate $rows[$i, 3])
            Start           = (ConvertFrom-TaskTableDate $rows[$i, 4])
            PercentComplete = $(if ($null -ne $rows[$i, 6]) { [int]([double]$rows[$i, 6] * 100) } else { 0 })
            Importance      = $script:ImportanceNames[[int]$rows[$i, 8]]
            Categories      = [string]$rows[$i, 9]
            Completed       = (ConvertFrom-TaskTableDate $rows[$i, 11])
            Modified        = $rows[$i, 10]
            SourceEntryID   = [string]$rows[$i, 13]
            Recurring       = [bool]$rows[$i, 14]
            Recurrence      = ""   # pattern text; filled per item by Add-TaskRecurrenceText (shown records only)
            Reminder        = $(if ($rows[$i, 15]) { ConvertFrom-TableUtc $rows[$i, 16] } else { $null })
            SourceMessageID = [string]$rows[$i, 17]
        }
        $List.Add($rec)   # not inline: Windows PowerShell 5.1 cannot parse a multi-line [ordered]@{} inside a method call
    }
}
# True when a record falls in a due window: Today, Overdue, Week (today + 6 days), NoDate.
# Completed items are never due or overdue.
function Test-TaskDue($Record, [string]$Due) {
    $today = (Get-Date).Date
    $d = $Record.Due
    if ($Due -in "Today", "Overdue", "Week" -and $Record.Status -eq "Completed") { return $false }
    switch ($Due) {
        "Today" { return [bool]($d -and $d.Date -eq $today) }
        "Overdue" { return [bool]($d -and $d.Date -lt $today) }
        "Week" { return [bool]($d -and $d.Date -ge $today -and $d.Date -le $today.AddDays(6)) }
        "NoDate" { return (-not $d) }
    }
    return $true
}

# The task an action script targets: -EntryID (preferred), -Index (the #N that
# outlook-tasks-list.ps1 prints when it runs without filters) or -Subject (exact; an open
# task wins over completed ones, and the match must be unique). Returns the TaskItem.
function Resolve-TaskItem {
    param($Namespace, [string]$EntryID = "", [int]$Index = 0, [string]$Subject = "", [switch]$IncludeCompleted, [switch]$AllowMail)
    if ($EntryID) {
        $item = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID
        $class = [int]$item.Class
        if ($class -eq 48) { return $item }                       # olTask
        if ($AllowMail -and $class -eq 43) { return $item }       # olMail (flagged mail as a to-do)
        throw "That EntryID is not a task (item class $class). Use outlook-tasks-list.ps1 to get task EntryIDs."
    }
    if ($Index -gt 0) {
        Write-Host "Note: -Index is volatile (positions shift as tasks change). Prefer -EntryID." -ForegroundColor DarkGray
        $recs = (Get-TaskRecords -Namespace $Namespace -IncludeCompleted:$IncludeCompleted).Records
        if ($Index -gt $recs.Count) { throw "No task at -Index $Index ($($recs.Count) tasks). Use -EntryID from outlook-tasks-list.ps1." }
        return (Get-OutlookItem -Namespace $Namespace -EntryID $recs[$Index - 1].EntryID)
    }
    if ($Subject) {
        $hits = @((Get-TaskRecords -Namespace $Namespace -IncludeCompleted).Records | Where-Object { $_.Subject -eq $Subject })
        if ($hits.Count -eq 0) { throw "No task with the exact subject '$Subject'." }
        $open = @($hits | Where-Object { $_.Status -ne "Completed" })
        if ($open.Count -ge 1) { $hits = $open }
        if ($hits.Count -gt 1) { throw "$($hits.Count) tasks have the exact subject '$Subject'. Use -EntryID from outlook-tasks-list.ps1." }
        return (Get-OutlookItem -Namespace $Namespace -EntryID $hits[0].EntryID)
    }
    throw "Provide -EntryID (preferred), -Index or -Subject to choose a task."
}

# Record for one TaskItem (same fields as Get-TaskRecords).
function ConvertTo-TaskRecord($Task) {
    $due = $null; $start = $null; $done = $null
    try { if (([datetime]$Task.DueDate).Year -lt 4500) { $due = [datetime]$Task.DueDate } } catch { }
    try { if (([datetime]$Task.StartDate).Year -lt 4500) { $start = [datetime]$Task.StartDate } } catch { }
    try { if ($Task.Complete -and ([datetime]$Task.DateCompleted).Year -lt 4500) { $done = [datetime]$Task.DateCompleted } } catch { }
    $rec = [ordered]@{
        EntryID         = $Task.EntryID
        Type            = "Task"
        Subject         = $Task.Subject
        Status          = $script:TaskStatusNames[[int]$Task.Status]
        Due             = $due
        Start           = $start
        PercentComplete = [int]$Task.PercentComplete
        Importance      = $script:ImportanceNames[[int]$Task.Importance]
        Categories      = $Task.Categories
        Completed       = $done
        Modified        = $Task.LastModificationTime
        SourceEntryID   = ""
        SourceMessageID = ""
        Recurring       = [bool]$Task.IsRecurring
        Recurrence      = ""
        Reminder        = $(if ($Task.ReminderSet) { [datetime]$Task.ReminderTime } else { $null })
    }
    try { $rec.SourceEntryID = [string]$Task.PropertyAccessor.GetProperty($script:TaskProp.SourceEntryID) } catch { }
    try { $rec.SourceMessageID = [string]$Task.PropertyAccessor.GetProperty($script:TaskProp.SourceMessageID) } catch { }
    if ($rec.Recurring) { $rec.Recurrence = Get-TaskRecurrenceText $Task }
    return $rec
}

$script:RecurrenceNames = @{ 0 = "Daily"; 1 = "Weekly"; 2 = "Monthly"; 3 = "Monthly"; 5 = "Yearly"; 6 = "Yearly" }

# "Weekly every 2 weeks on Mon, Wed until 2026-12-31" for a recurring task ("" when not recurring).
function Get-TaskRecurrenceText($Task) {
    if (-not $Task.IsRecurring) { return "" }
    try {
        $p = $Task.GetRecurrencePattern()
        $type = [int]$p.RecurrenceType
        $name = $script:RecurrenceNames[$type]
        $unit = @{ 0 = "day"; 1 = "week"; 2 = "month"; 3 = "month"; 5 = "year"; 6 = "year" }[$type]
        $n = [int]$p.Interval
        if ($type -in 5, 6 -and $n -ge 12) { $n = [int]($n / 12) }   # yearly intervals are stored in months
        $text = if ($n -gt 1) { "$name every $n ${unit}s" } else { $name }
        if ($type -eq 1) {
            $days = @(); $mask = [int]$p.DayOfWeekMask
            $i = 0; foreach ($d in "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat") { if ($mask -band (1 -shl $i)) { $days += $d }; $i++ }
            if ($days) { $text += " on $($days -join ', ')" }
        }
        if ($type -eq 2) { $text += " on day $($p.DayOfMonth)" }
        if (-not $p.NoEndDate) {
            if ([int]$p.Occurrences -gt 0 -and ([datetime]$p.PatternEndDate).Year -ge 4500) { $text += ", $($p.Occurrences) times" }
            else { $text += " until $(([datetime]$p.PatternEndDate).ToString('yyyy-MM-dd'))" }
        }
        return $text
    } catch { return "recurring" }
}

# Fill Recurrence on the recurring records only (one item open each), e.g. for the records a list shows.
function Add-TaskRecurrenceText($Namespace, $Records) {
    foreach ($r in $Records) {
        if (-not $r.Recurring -or $r.Type -ne "Task") { continue }
        try { $r.Recurrence = Get-TaskRecurrenceText ($Namespace.GetItemFromID($r.EntryID)) } catch { $r.Recurrence = "recurring" }
    }
}

# JSON copy of a task record (dates as ISO, task dates as yyyy-MM-dd).
function ConvertTo-TaskJson($Record) {
    $o = [ordered]@{}
    foreach ($k in $Record.Keys) { $o[$k] = $Record[$k] }
    $o.Due = $(if ($Record.Due) { Format-TaskDate $Record.Due } else { $null })
    $o.Start = $(if ($Record.Start) { Format-TaskDate $Record.Start } else { $null })
    $o.Completed = $(if ($Record.Completed) { Format-TaskDate $Record.Completed } else { $null })
    $o.Modified = Format-MailDateIso $Record.Modified
    $o.Reminder = $(if ($Record.Reminder) { Format-MailDateIso $Record.Reminder } else { $null })
    if (-not $o.SourceEntryID) { $o.Remove("SourceEntryID") }
    if (-not $o.SourceMessageID) { $o.Remove("SourceMessageID") }
    if (-not $o.Recurrence) { $o.Remove("Recurrence") }
    return $o
}

# "Due: Thu, Oct 22, 2026 (TODAY)" with the matching color.
function Write-TaskDueLine($Due, [string]$Status) {
    if (-not $Due) { return }
    $today = (Get-Date).Date
    $label = ""; $color = "Gray"
    if ($Status -ne "Completed") {
        if ($Due.Date -lt $today) { $label = " (OVERDUE)"; $color = "Red" }
        elseif ($Due.Date -eq $today) { $label = " (TODAY)"; $color = "Yellow" }
        elseif ($Due.Date -eq $today.AddDays(1)) { $label = " (Tomorrow)" }
    }
    Write-Host "  Due: $($Due.ToString('ddd, MMM d, yyyy'))$label" -ForegroundColor $color
}

# Text rendering of one task record.
function Write-TaskRecord($Record, [string]$Heading = "") {
    if ($Heading) { Write-Host $Heading -ForegroundColor Cyan }
    Write-Host "  EntryID: $($Record.EntryID)"
    Write-Host "  Subject: $($Record.Subject)" -ForegroundColor Yellow
    if ($Record.Type -eq "Mail") { Write-Host "  Type: flagged email" -ForegroundColor Magenta }
    $line = "  Status: $($script:TaskStatusLabels[$script:TaskStatusMap[$Record.Status]])"
    if ($Record.Status -eq "InProgress" -and $Record.PercentComplete -gt 0) { $line += " ($($Record.PercentComplete)%)" }
    if ($Record.Importance -and $Record.Importance -ne "Normal") { $line += " | Priority: $($Record.Importance)" }
    Write-Host $line -ForegroundColor Gray
    Write-TaskDueLine $Record.Due $Record.Status
    if ($Record.Start) { Write-Host "  Start: $($Record.Start.ToString('ddd, MMM d, yyyy'))" -ForegroundColor Gray }
    if ($Record.Recurring) { Write-Host "  Repeats: $(if ($Record.Recurrence) { $Record.Recurrence } else { 'yes' })" -ForegroundColor Magenta }
    if ($Record.Reminder) { Write-Host "  Reminder: $(Format-MailDate $Record.Reminder)" -ForegroundColor Gray }
    if ($Record.Completed) { Write-Host "  Completed: $($Record.Completed.ToString('ddd, MMM d, yyyy'))" -ForegroundColor Green }
    if ($Record.Categories) { Write-Host "  Categories: $($Record.Categories)" -ForegroundColor Green }
    if ($Record.SourceEntryID) { Write-Host "  From email: $($Record.SourceEntryID)" -ForegroundColor DarkGray }
}

# Result of an action on one task: title, the task, what changed. -Json prints one object.
function Write-TaskResult {
    param([string]$Title, $Task, $Changes = [ordered]@{}, [switch]$Json, [string]$Action = "", $Record = $null)
    $rec = if ($Record) { $Record } else { ConvertTo-TaskRecord $Task }
    if ($Json) {
        $o = ConvertTo-TaskJson $rec
        $o.Action = $Action; $o.Result = $Title
        # Nested: change labels (Due, Status, ...) would otherwise overwrite the task's own fields.
        if ($Changes.Count -gt 0) { $o.Changes = $Changes }
        Write-Json $o
        return
    }
    Write-TaskRecord $rec "=== $Title ==="
    foreach ($k in $Changes.Keys) { Write-Host "  ${k}: $($Changes[$k])" -ForegroundColor Cyan }
}

# Preview of an irreversible task action, then the confirm gate. Returns $true to proceed.
# The ONLY output is that bool: anything else written to the pipeline here would turn the
# result into an array, which PowerShell treats as true, and the action would run unconfirmed.
# With -Json the preview is kept in $script:TaskPreviewJson for the caller to print.
function Confirm-TaskAction {
    param([string]$Action, [string]$Title, $Task, [bool]$Confirmed, $Details = [ordered]@{}, [string]$What = "", [switch]$Json)
    $script:TaskPreviewJson = $null
    if ($Confirmed -or -not (Test-ConfirmRequired -Action $Action)) { return $true }
    if ($Json) {
        $o = ConvertTo-TaskJson (ConvertTo-TaskRecord $Task)
        $o.Action = $Action; $o.Result = "preview"; $o.Done = $false
        if ($Details.Count -gt 0) { $o.Details = $Details }
        $o.Hint = "Nothing was done. Run the same command with -Confirm to do it."
        $script:TaskPreviewJson = $o
        return $false
    }
    Write-TaskRecord (ConvertTo-TaskRecord $Task) "=== $Title (PREVIEW) ==="
    foreach ($k in $Details.Keys) { Write-Host "  ${k}: $($Details[$k])" -ForegroundColor Cyan }
    return (Approve-Action -Action $Action -Confirmed $false -What $(if ($What) { $What } else { $Title.ToLower() }))
}
