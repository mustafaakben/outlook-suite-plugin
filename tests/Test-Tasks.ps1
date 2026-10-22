<#
.SYNOPSIS
  Live test for the outlook-tasks scripts, on disposable "[suite-test]" tasks only.

.DESCRIPTION
  Creates test tasks, runs every tasks script against them, checks the real Outlook state
  after each step, and prints PASS/FAIL lines and counts only. Nothing is sent: assignment
  is run in preview mode only (the confirm gate must block it). A real flagged email is read
  for tasks-from-mail but never modified. Cleans up the test tasks at the end.
  Run under both shells:  pwsh -File Test-Tasks.ps1   /   powershell -File Test-Tasks.ps1 -Shell powershell
#>
param(
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-tasks/scripts"),
    [ValidateSet("pwsh", "powershell")]
    [string]$Shell = "pwsh"
)

$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
. (Join-Path $ScriptsDir "_tasks.ps1")
$null = Connect-Outlook
# Guard: nothing may be sent. Sent Items and Outbox counts are compared at the end; any increase
# fails the run (a send by someone else during the run also trips it: re-run to confirm).
$sentBefore = $Namespace.GetDefaultFolder(5).Items.Count; $outboxBefore = $Namespace.GetDefaultFolder(4).Items.Count
$savedGate = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"   # force the gate on: a user config that turns it off must never let a test send or delete

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $ErrorActionPreference = "Continue"   # Windows PowerShell 5.1 turns child stderr into terminating errors under Stop
    $out = & $Shell -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}
function RunJson([string]$script, [object[]]$a) {
    $r = Run $script ($a + "-Json")
    try { return [pscustomobject]@{ Exit = $r.Exit; Obj = ($r.Out | ConvertFrom-Json); Out = $r.Out } } catch { return [pscustomobject]@{ Exit = $r.Exit; Obj = $null; Out = $r.Out } }
}
function Item($id) { return $Namespace.GetItemFromID($id) }

$stamp = Get-Date -Format "HHmmss"
$created = New-Object System.Collections.Generic.List[string]
$today = (Get-Date).Date

try {
    # --- create ---
    $r = RunJson "outlook-tasks-create.ps1" @("-Subject", "[suite-test] due today $stamp", "-DueDate", $today.ToString("yyyy-MM-dd"), "-Priority", "High", "-Categories", "SuiteTest")
    $idToday = $r.Obj.EntryID; if ($idToday) { $created.Add($idToday) }
    Check "create (json, due today, high)" ($r.Exit -eq 0 -and $idToday -and $r.Obj.Due -eq $today.ToString("yyyy-MM-dd") -and $r.Obj.Importance -eq "High") "exit=$($r.Exit)"
    $r = RunJson "outlook-tasks-create.ps1" @("-Subject", "[suite-test] overdue $stamp", "-DueDate", $today.AddDays(-3).ToString("yyyy-MM-dd"))
    $idOver = $r.Obj.EntryID; if ($idOver) { $created.Add($idOver) }
    Check "create (overdue)" ($r.Exit -eq 0 -and $idOver)
    $r = Run "outlook-tasks-create.ps1" @("-Subject", "[suite-test] no date $stamp", "-Reminder", $today.AddDays(5).AddHours(9).ToString("yyyy-MM-dd HH:mm"))
    $idNo = (Get-TaskRecords -Namespace $Namespace).Records | Where-Object { $_.Subject -eq "[suite-test] no date $stamp" } | Select-Object -First 1 | ForEach-Object { $_.EntryID }
    if ($idNo) { $created.Add($idNo) }
    Check "create (text output, reminder)" ($r.Exit -eq 0 -and $r.Out -match "TASK CREATED" -and $idNo -and (Item $idNo).ReminderSet)

    # --- list views ---
    $l = RunJson "outlook-tasks-list.ps1" @("-Due", "Today", "-Search", "[suite-test]", "-Limit", "50")
    Check "list -Due Today finds the today task only" ($l.Exit -eq 0 -and @($l.Obj.items | Where-Object { $_.EntryID -eq $idToday }).Count -eq 1 -and @($l.Obj.items | Where-Object { $_.EntryID -in $idOver, $idNo }).Count -eq 0)
    $l = RunJson "outlook-tasks-list.ps1" @("-Due", "Overdue", "-Search", "suite-test] overdue $stamp")
    Check "list -Due Overdue" ($l.Exit -eq 0 -and $l.Obj.total -eq 1 -and $l.Obj.items[0].EntryID -eq $idOver)
    $l = RunJson "outlook-tasks-list.ps1" @("-Due", "Week", "-Search", "[suite-test]", "-Limit", "50")
    Check "list -Due Week includes today, not overdue" ($l.Exit -eq 0 -and @($l.Obj.items.EntryID) -contains $idToday -and -not (@($l.Obj.items.EntryID) -contains $idOver))
    $l = RunJson "outlook-tasks-list.ps1" @("-Due", "NoDate", "-Search", "[suite-test] no date $stamp")
    Check "list -Due NoDate" ($l.Exit -eq 0 -and $l.Obj.total -eq 1)
    $l = RunJson "outlook-tasks-list.ps1" @("-Category", "SuiteTest", "-Search", $stamp)
    Check "list -Category" ($l.Exit -eq 0 -and $l.Obj.total -eq 1 -and $l.Obj.items[0].EntryID -eq $idToday)
    # The To-Do List is a search folder that Outlook updates asynchronously: allow a few seconds.
    for ($try = 0; $try -lt 10; $try++) {
        $l = RunJson "outlook-tasks-list.ps1" @("-Flagged", "-Search", "[suite-test]", "-Due", "Today", "-Limit", "50")
        if (@($l.Obj.items.EntryID) -contains $idToday) { break }
        Start-Sleep -Seconds 1
    }
    Check "list -Flagged includes tasks" ($l.Exit -eq 0 -and @($l.Obj.items.EntryID) -contains $idToday)
    $all = RunJson "outlook-tasks-list.ps1" @("-Flagged", "-Limit", "5000")
    Check "list -Flagged includes flagged email" ($all.Exit -eq 0 -and @($all.Obj.items | Where-Object { $_.Type -eq "Mail" }).Count -gt 0) "mail=$(@($all.Obj.items | Where-Object { $_.Type -eq 'Mail' }).Count)"
    $t = Run "outlook-tasks-list.ps1" @("-Due", "Today", "-Search", $stamp)
    Check "list text output" ($t.Exit -eq 0 -and $t.Out -match "OUTLOOK TASKS" -and $t.Out -match "\(TODAY\)")

    # Independent cross-check of the open-task count (Items.Restrict, no table).
    $open = $Namespace.GetDefaultFolder(13).Items.Restrict("[Complete] = False")
    $openTasks = 0; foreach ($i in $open) { if ($i.MessageClass -like "IPM.Task*") { $openTasks++ } }
    $l = RunJson "outlook-tasks-list.ps1" @("-Limit", "1")
    Check "open-task count matches Items.Restrict ($openTasks)" ($l.Obj.total -eq $openTasks) "table=$($l.Obj.total)"

    # --- update ---
    $r = RunJson "outlook-tasks-update.ps1" @("-EntryID", $idToday, "-PercentComplete", "40", "-Subject", "[suite-test] renamed $stamp", "-Categories", "SuiteTest2")
    $it = Item $idToday
    Check "update progress auto-sets InProgress, rename, categories" ($r.Exit -eq 0 -and $it.PercentComplete -eq 40 -and $it.Status -eq 1 -and $it.Subject -eq "[suite-test] renamed $stamp" -and $it.Categories -eq "SuiteTest2")
    $r = Run "outlook-tasks-update.ps1" @("-EntryID", $idToday, "-Status", "Completed", "-PercentComplete", "50")
    Check "update rejects status/percent conflict" ($r.Exit -eq 1 -and (Item $idToday).Status -eq 1)
    $r = Run "outlook-tasks-update.ps1" @("-EntryID", $idToday)
    Check "update with no changes fails" ($r.Exit -eq 1)
    $r = RunJson "outlook-tasks-update.ps1" @("-EntryID", $idOver, "-ClearDueDate")
    Check "update -ClearDueDate" ($r.Exit -eq 0 -and -not $r.Obj.Due -and ([datetime](Item $idOver).DueDate).Year -ge 4500)

    # --- complete ---
    $r = RunJson "outlook-tasks-complete.ps1" @("-Subject", "[suite-test] renamed $stamp")
    Check "complete by exact subject" ($r.Exit -eq 0 -and (Item $idToday).Complete)
    $r = Run "outlook-tasks-complete.ps1" @("-EntryID", $idToday)
    Check "complete again says already completed" ($r.Exit -eq 0 -and $r.Out -match "ALREADY")
    $r = Run "outlook-tasks-complete.ps1" @("-Subject", "[suite-test] does not exist $stamp")
    Check "complete unknown subject fails" ($r.Exit -eq 1)

    # --- assign: preview only, nothing may be sent ---
    $r = Run "outlook-tasks-assign.ps1" @("-EntryID", $idNo, "-AssignTo", "nobody@example.invalid")
    $it = Item $idNo
    Check "assign without -Confirm is a preview (gate blocks it)" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY" -and -not $it.Delegator -and [int]$it.DelegationState -eq 0)
    $r = Run "outlook-tasks-assign.ps1" @("-EntryID", $idNo, "-AssignTo", "zz-no-such-person-$stamp@@invalid")
    Check "assign to an unresolvable name fails before any preview" ($r.Exit -eq 1)

    # --- confirm gate: the helper must return exactly [bool] $false when not confirmed, in text and -Json mode ---
    foreach ($mode in $false, $true) {
        $g = Confirm-TaskAction -Action "delete" -Title "GATE CHECK" -Task (Item $idNo) -Confirmed $false -Json:$mode 6>$null
        Check "Confirm-TaskAction returns [bool] `$false (Json=$mode)" ($g -is [bool] -and $g -eq $false) "type=$(if ($null -eq $g) { 'null' } else { $g.GetType().Name })"
    }
    $p = RunJson "outlook-tasks-delete.ps1" @("-EntryID", $idNo)
    Check "delete -Json without -Confirm prints a preview and deletes nothing" ($p.Exit -eq 0 -and $p.Obj.Result -eq "preview" -and $p.Obj.Done -eq $false -and (Item $idNo).Subject)
    $p = RunJson "outlook-tasks-assign.ps1" @("-EntryID", $idNo, "-AssignTo", "nobody@example.invalid")
    Check "assign -Json without -Confirm is a JSON preview or a resolve error, never a send" ($p.Exit -in 0, 1 -and -not (Item $idNo).Delegator)

    # --- delete: preview, then confirmed ---
    $r = Run "outlook-tasks-delete.ps1" @("-EntryID", $idOver)
    Check "delete without -Confirm is a preview" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY" -and (Item $idOver))
    $r = RunJson "outlook-tasks-delete.ps1" @("-EntryID", $idOver, "-Confirm")
    # Check the folders, not GetItemFromID: this session may still hold the old item in its cache.
    $inDeleted = @($Namespace.GetDefaultFolder(3).Items.Restrict("[Subject] = '[suite-test] overdue $stamp'")).Count
    $inTasks = @((Get-TaskRecords -Namespace $Namespace).Records | Where-Object { $_.Subject -eq "[suite-test] overdue $stamp" }).Count
    $gone = ($inDeleted -eq 1 -and $inTasks -eq 0)
    Check "delete -Confirm moves to Deleted Items" ($r.Exit -eq 0 -and $r.Obj.Result -eq "deleted" -and $gone)
    if ($gone) { [void]$created.Remove($idOver) }

    # --- recurring tasks and reminders ---
    $due = $today.AddDays(7)
    $until = $due.AddDays(42)
    $r = RunJson "outlook-tasks-create.ps1" @("-Subject", "[suite-test] weekly $stamp", "-Recurrence", "Weekly", "-Interval", "2", "-DueDate", $due.ToString("yyyy-MM-dd"), "-Until", $until.ToString("yyyy-MM-dd"), "-Reminder", $due.AddHours(9).ToString("yyyy-MM-dd HH:mm"))
    $idWeekly = $r.Obj.EntryID; if ($idWeekly) { $created.Add($idWeekly) }
    $it = $null; $p = $null; if ($idWeekly) { $it = Item $idWeekly; $p = $it.GetRecurrencePattern() }
    Check "create -Recurrence Weekly -Interval 2 -Until -Reminder" ($r.Exit -eq 0 -and $it.IsRecurring -and $p.RecurrenceType -eq 1 -and $p.Interval -eq 2 -and $p.DayOfWeekMask -eq (1 -shl [int]$due.DayOfWeek) -and ([datetime]$p.PatternEndDate).Date -eq $until -and $it.ReminderSet) "exit=$($r.Exit)"
    $r = RunJson "outlook-tasks-create.ps1" @("-Subject", "[suite-test] monthly $stamp", "-Recurrence", "Monthly", "-DueDate", $due.ToString("yyyy-MM-dd"))
    $idMonthly = $r.Obj.EntryID; if ($idMonthly) { $created.Add($idMonthly) }
    $p = $null; if ($idMonthly) { $p = (Item $idMonthly).GetRecurrencePattern() }
    Check "create -Recurrence Monthly (no end)" ($r.Exit -eq 0 -and $p.RecurrenceType -eq 2 -and $p.DayOfMonth -eq $due.Day -and $p.NoEndDate)
    $r = RunJson "outlook-tasks-create.ps1" @("-Subject", "[suite-test] daily $stamp", "-Recurrence", "Daily")
    $idDaily = $r.Obj.EntryID; if ($idDaily) { $created.Add($idDaily) }
    Check "create -Recurrence Daily defaults the due date to today" ($r.Exit -eq 0 -and $r.Obj.Due -eq $today.ToString("yyyy-MM-dd") -and (Item $idDaily).IsRecurring)
    $r = Run "outlook-tasks-create.ps1" @("-Subject", "[suite-test] bad $stamp", "-Until", $until.ToString("yyyy-MM-dd"))
    Check "create -Until without -Recurrence fails" ($r.Exit -eq 1)
    $l = RunJson "outlook-tasks-list.ps1" @("-Search", "weekly $stamp")
    Check "list shows recurrence and reminder (json)" ($l.Exit -eq 0 -and $l.Obj.items[0].Recurring -and $l.Obj.items[0].Recurrence -match "^Weekly every 2 weeks on \w{3} until" -and $l.Obj.items[0].Reminder)
    $t = Run "outlook-tasks-list.ps1" @("-Search", "monthly $stamp")
    Check "list shows recurrence (text)" ($t.Exit -eq 0 -and $t.Out -match "Repeats: Monthly on day $($due.Day)")

    # --- date-only dates across the Nov 1 DST change ---
    # Classic Outlook saves local midnight; the checks below run through the scripts.
    foreach ($day in "2026-11-01", "2026-11-05") {
        $r = RunJson "outlook-tasks-create.ps1" @("-Subject", "[suite-test] due $day $stamp", "-DueDate", $day)
        if ($r.Obj.EntryID) { $created.Add($r.Obj.EntryID) }
        $l = RunJson "outlook-tasks-list.ps1" @("-Search", "due $day $stamp")
        Check "due $day (classic, local midnight) lists as $day" ($l.Obj.items[0].Due -eq $day) "got $($l.Obj.items[0].Due)"
    }
    # Outlook on the web, mobile, To Do and Graph save 00:00 UTC. Classic Outlook rewrites a task's
    # due date to local midnight on Save, so such an item cannot be made here: check the conversion.
    $cases = @(
        @("2026-11-01 00:00", "2026-11-01", "web 00:00 UTC on the DST day"),
        @("2026-11-05 00:00", "2026-11-05", "web 00:00 UTC after DST"),
        @("2026-10-22 00:00", "2026-10-22", "web 00:00 UTC before DST")
    )
    if ([TimeZoneInfo]::Local.Id -eq "Eastern Standard Time") {
        $cases += , @("2026-11-01 04:00", "2026-11-01", "classic local midnight on the DST day (EDT, 04:00 UTC)")
        $cases += , @("2026-11-05 05:00", "2026-11-05", "classic local midnight after DST (EST, 05:00 UTC)")
    }
    foreach ($k in $cases) {
        $got = ConvertFrom-TaskTableDate ([datetime]$k[0])
        Check "date conversion: $($k[2])" ($got -eq [datetime]$k[1] -and $got.TimeOfDay.Ticks -eq 0) "got $got"
    }
    $rem = ConvertFrom-TableUtc ([datetime]"2026-11-05 00:00")
    Check "reminders are real times (00:00 UTC stays the previous evening locally)" ($rem -eq [datetime]::SpecifyKind([datetime]"2026-11-05", [DateTimeKind]::Utc).ToLocalTime())
    # --- review fixes: -Index numbering, open-first -Subject, JSON previews, categories ---
    $t = Run "outlook-tasks-list.ps1" @("-Limit", "1")
    $tf = Run "outlook-tasks-list.ps1" @("-Search", $stamp, "-Limit", "1")
    Check "#N printed only for the unfiltered list" ($t.Out -match "--- #1 ---" -and $tf.Out -notmatch "--- #1 ---" -and $tf.Out -match "--- result 1 ---")
    $twinA = RunJson "outlook-tasks-create.ps1" @("-Subject", "[suite-test] twin $stamp"); $created.Add($twinA.Obj.EntryID)
    $twinB = RunJson "outlook-tasks-create.ps1" @("-Subject", "[suite-test] twin $stamp"); $created.Add($twinB.Obj.EntryID)
    $null = Run "outlook-tasks-complete.ps1" @("-EntryID", $twinA.Obj.EntryID)
    $r = RunJson "outlook-tasks-complete.ps1" @("-Subject", "[suite-test] twin $stamp")
    Check "complete -Subject prefers the open task over a completed twin" ($r.Exit -eq 0 -and $r.Obj.EntryID -eq $twinB.Obj.EntryID -and (Item $twinB.Obj.EntryID).Complete)
    $l = RunJson "outlook-tasks-list.ps1" @("-Due", "NoDate", "-IncludeCompleted", "-Search", "twin $stamp")
    $lw = RunJson "outlook-tasks-list.ps1" @("-Due", "Today", "-IncludeCompleted", "-Search", "renamed $stamp")
    Check "completed tasks are never due or overdue" ($l.Obj.total -eq 2 -and $lw.Obj.total -eq 0)
    $p = RunJson "outlook-tasks-delete.ps1" @("-EntryID", $twinA.Obj.EntryID)
    Check "delete preview with -Json prints JSON" ($p.Exit -eq 0 -and $p.Obj.Result -eq "preview" -and -not $p.Obj.Done -and (Item $twinA.Obj.EntryID))
    $p = RunJson "outlook-tasks-assign.ps1" @("-EntryID", $twinA.Obj.EntryID, "-AssignTo", "nobody@example.invalid")
    Check "assign preview with -Json prints JSON (nothing sent)" ($p.Exit -eq 0 -and $p.Obj.Result -eq "preview" -and $p.Obj.Details)
    $null = Run "outlook-tasks-update.ps1" @("-EntryID", $twinA.Obj.EntryID, "-Categories", "SuiteTestA; SuiteTestB", "-IncludeCompleted")
    $l = RunJson "outlook-tasks-list.ps1" @("-Category", "SuiteTestB", "-Search", "twin $stamp", "-IncludeCompleted")
    Check "-Category matches one of several (split on , and ;)" ($l.Obj.total -eq 1)

    # --- task from email (the email is only read: its state must not change) ---
    $inbox = $Namespace.GetDefaultFolder(6)
    $tb = $inbox.GetTable("[MessageClass] = 'IPM.Note'"); $tb.Sort("ReceivedTime", $true)
    $mail = Item ($tb.GetNextRow().Item("EntryID"))
    $mailId = $mail.EntryID   # the item's own (long-term) EntryID, which the task stores
    $before = "$($mail.LastModificationTime.ToString('o'))|$($mail.FlagStatus)|$($mail.UnRead)|$($mail.Categories)"
    $r = RunJson "outlook-tasks-from-mail.ps1" @("-EntryID", $mailId, "-Subject", "[suite-test] from mail $stamp", "-When", "Tomorrow", "-AttachEmail")
    $idMail = $r.Obj.EntryID; if ($idMail) { $created.Add($idMail) }
    $it = $null; if ($idMail) { $it = Item $idMail }
    Check "from-mail creates a linked task" ($r.Exit -eq 0 -and $it -and $it.Body -match [regex]::Escape("outlook:$mailId") -and $it.Attachments.Count -eq 1 -and $r.Obj.Due -eq $today.AddDays(1).ToString("yyyy-MM-dd")) "exit=$($r.Exit)"
    $l = RunJson "outlook-tasks-list.ps1" @("-Search", "from mail $stamp")
    Check "list shows the source email" ($l.Exit -eq 0 -and $l.Obj.items[0].SourceEntryID -eq $mailId)
    $msgId = ""; try { $msgId = [string]$mail.PropertyAccessor.GetProperty($script:TaskProp.InternetMessageID) } catch { }
    Check "task stores the email's Message-ID (survives moves)" ($msgId -and $l.Obj.items[0].SourceMessageID -eq $msgId)
    $r = Run "outlook-tasks-from-mail.ps1" @("-EntryID", $mailId)
    Check "from-mail refuses a duplicate" ($r.Exit -eq 1 -and $r.Out -match "already exists")
    $r = RunJson "outlook-tasks-from-mail.ps1" @("-EntryID", $mailId, "-Subject", "[suite-test] from mail again $stamp", "-Force")
    if ($r.Obj.EntryID) { $created.Add($r.Obj.EntryID) }
    Check "from-mail -Force allows a second task" ($r.Exit -eq 0 -and $r.Obj.EntryID)
    $past = $today.AddDays(-2)
    $r = RunJson "outlook-tasks-from-mail.ps1" @("-EntryID", $mailId, "-Subject", "[suite-test] from mail past $stamp", "-DueDate", $past.ToString("yyyy-MM-dd"), "-Force")
    if ($r.Obj.EntryID) { $created.Add($r.Obj.EntryID) }
    Check "from-mail with a past due date starts on the due date, not after it" ($r.Exit -eq 0 -and $r.Obj.Due -eq $past.ToString("yyyy-MM-dd") -and $r.Obj.Start -eq $past.ToString("yyyy-MM-dd"))    $mail = $Namespace.GetItemFromID($mailId)
    $after = "$($mail.LastModificationTime.ToString('o'))|$($mail.FlagStatus)|$($mail.UnRead)|$($mail.Categories)"
    Check "source email unchanged" ($before -eq $after)
    $r = Run "outlook-tasks-from-mail.ps1" @("-EntryID", $idToday)
    Check "from-mail rejects a non-email EntryID" ($r.Exit -eq 1)

    # --- errors ---
    $r = Run "outlook-tasks-update.ps1" @("-EntryID", "00", "-Subject", "x")
    Check "bad EntryID fails with exit 1" ($r.Exit -eq 1 -and $r.Out -match "Error")
}
finally {
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $savedGate
    # Cleanup, leaving nothing behind: move each test task to Deleted Items and delete it there.
    # Deleting an object this session opened, moved or re-opened does not reliably purge it,
    # so: delete (moves to Deleted Items), then find each subject there and delete that copy.
    $del = $Namespace.GetDefaultFolder(3)
    $subjects = New-Object System.Collections.Generic.List[string]
    $subjects.Add("[suite-test] overdue $stamp")
    foreach ($id in $created) { try { $x = Item $id; $subjects.Add($x.Subject); $x.Delete() } catch { } }
    $n = 0
    for ($try = 0; $try -lt 10 -and $subjects.Count -gt 0; $try++) {
        foreach ($s in @($subjects)) {
            $hits = @($del.Items.Restrict("[Subject] = '$($s.Replace("'", "''"))'"))
            foreach ($x in $hits) { try { $x.Delete(); $n++ } catch { } }
            if ($hits.Count -gt 0) { [void]$subjects.Remove($s) }
        }
        if ($subjects.Count -gt 0) { Start-Sleep -Seconds 1 }
    }
    Write-Host "Cleanup: $n test item(s) removed$(if ($subjects.Count) { "; not found in Deleted Items: $($subjects.Count)" })." -ForegroundColor Gray
}
$sentAfter = $Namespace.GetDefaultFolder(5).Items.Count; $outboxAfter = $Namespace.GetDefaultFolder(4).Items.Count
Check "nothing sent (Sent Items $sentBefore -> $sentAfter, Outbox $outboxBefore -> $outboxAfter)" ($sentAfter -le $sentBefore -and $outboxAfter -le $outboxBefore)
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASS ($Shell)" -ForegroundColor Green } else { Write-Host "$($script:fails) FAILED ($Shell)" -ForegroundColor Red }
exit $script:fails
