<#
.SYNOPSIS
  Live test for outlook-tasks-triage.ps1, on disposable "[suite-test]" tasks only.

.DESCRIPTION
  Creates four overdue test tasks (3, 10 and 40 days overdue, plus a recurring one), checks the
  grouping, the previews (nothing changes without -Confirm) and each bulk action, always
  narrowed with -Search to this run's tasks. Checks that the number of real overdue tasks is
  the same at the end, that nothing was sent, and removes the test tasks. Prints PASS/FAIL and
  counts only. Run under both shells:  pwsh -File Test-TasksTriage.ps1  /  powershell -File Test-TasksTriage.ps1 -Shell powershell
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
$sentBefore = $Namespace.GetDefaultFolder(5).Items.Count; $outboxBefore = $Namespace.GetDefaultFolder(4).Items.Count
$savedGate = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"   # force the gate on

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $ErrorActionPreference = "Continue"
    $out = & $Shell -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}
function RunJson([string]$script, [object[]]$a) {
    $r = Run $script ($a + "-Json")
    try { return [pscustomobject]@{ Exit = $r.Exit; Obj = ($r.Out | ConvertFrom-Json); Out = $r.Out } } catch { return [pscustomobject]@{ Exit = $r.Exit; Obj = $null; Out = $r.Out } }
}
function Item($id) { return $Namespace.GetItemFromID($id) }
function Due($id) { $d = [datetime](Item $id).DueDate; if ($d.Year -ge 4500) { return $null }; return $d.Date }
function RealOverdue { return @((Get-TaskRecords -Namespace $Namespace).Records | Where-Object { (Test-TaskDue $_ "Overdue") -and $_.Subject -notlike "*suite-test*" }).Count }

$stamp = Get-Date -Format "HHmmss"
$today = (Get-Date).Date
$tag = "triage $stamp"
$created = New-Object System.Collections.Generic.List[string]
$realBefore = RealOverdue

try {
    $specs = @(
        @("[suite-test] $tag A", 3, @("-Categories", "SuiteTriage")),
        @("[suite-test] $tag B", 10, @()),
        @("[suite-test] $tag C", 40, @("-Categories", "SuiteTriage")),
        @("[suite-test] $tag D", 14, @("-Recurrence", "Weekly")),
        @("[suite-test] $tag E", 5, @()),
        @("[suite-test] $tag G", 20, @())
    )
    $ids = @{}
    foreach ($s in $specs) {
        $r = RunJson "outlook-tasks-create.ps1" (@("-Subject", $s[0], "-DueDate", $today.AddDays(-$s[1]).ToString("yyyy-MM-dd")) + $s[2])
        $ids[$s[0].Substring($s[0].Length - 1)] = $r.Obj.EntryID
        if ($r.Obj.EntryID) { $created.Add($r.Obj.EntryID) }
    }
    # E is parked (Waiting). G reports completion to someone (an undeliverable test address), like an assigned task.
    $null = Run "outlook-tasks-update.ps1" @("-EntryID", $ids.E, "-Status", "Waiting")
    $g = Item $ids.G; $g.StatusOnCompletionRecipients = "nobody@example.invalid"; $g.Save()
    Check "six overdue test tasks created (one Waiting, one reporting completion)" ($created.Count -eq 6 -and (Item $ids.E).Status -eq 3 -and (Item $ids.G).StatusOnCompletionRecipients)

    # Report
    $rep = RunJson "outlook-tasks-triage.ps1" @("-Search", $tag)
    Check "report groups by age (1 / 3 / 1), Waiting left out" ($rep.Exit -eq 0 -and $rep.Obj.matched -eq 5 -and $rep.Obj.byAge."Under 1 week" -eq 1 -and $rep.Obj.byAge."1 to 4 weeks" -eq 3 -and $rep.Obj.byAge."Over a month" -eq 1 -and $rep.Obj.parkedExcluded -ge 1) "exit=$($rep.Exit) matched=$($rep.Obj.matched)"
    $cat = @($rep.Obj.byCategory | Where-Object { $_.Category -eq "SuiteTriage" })
    Check "report groups by category" ($cat.Count -eq 1 -and $cat[0].Count -eq 2 -and @($rep.Obj.byCategory | Where-Object { $_.Category -eq "(none)" })[0].Count -eq 3)
    $rep = RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, "-IncludeParked")
    Check "-IncludeParked brings the Waiting task back" ($rep.Obj.matched -eq 6)
    $t = Run "outlook-tasks-triage.ps1" @("-Search", $tag)
    Check "report text output" ($t.Exit -eq 0 -and $t.Out -match "OVERDUE TASKS: 5" -and $t.Out -match "By category" -and $t.Out -match "Waiting/Deferred left out")
    $o = RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, "-Older", "30")
    Check "-Older 30 keeps only the 40-day task" ($o.Obj.matched -eq 1)

    # Guards
    $r = Run "outlook-tasks-triage.ps1" @("-ClearDue")
    Check "an action without a filter is refused" ($r.Exit -eq 1 -and $r.Out -match "needs a filter")
    $r = Run "outlook-tasks-triage.ps1" @("-Search", $tag, "-Reschedule", $today.AddDays(-1).ToString("yyyy-MM-dd"))
    Check "reschedule into the past is refused" ($r.Exit -eq 1)
    $r = Run "outlook-tasks-triage.ps1" @("-Search", $tag, "-Complete", "-ClearDue")
    Check "two actions at once are refused" ($r.Exit -eq 1)

    # Previews change nothing and skip recurring and assigned tasks for every action
    foreach ($act in "-ClearDue", "-Complete") {
        $p = RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, $act)
        Check "preview $act (json): 3 would change, 1 recurring and 1 assigned skipped, plan code" ($p.Exit -eq 0 -and $p.Obj.result -eq "preview" -and $p.Obj.wouldChange -eq 3 -and $p.Obj.skippedRecurring -eq 1 -and $p.Obj.skippedAssigned -eq 1 -and $p.Obj.plan -match '^[0-9a-f]{12}$') "would=$($p.Obj.wouldChange) rec=$($p.Obj.skippedRecurring) asg=$($p.Obj.skippedAssigned)"
    }
    $p = Run "outlook-tasks-triage.ps1" @("-Search", $tag, "-Complete")
    Check "preview (text) says PREVIEW ONLY with a plan, completes nothing" ($p.Exit -eq 0 -and $p.Out -match "PREVIEW ONLY" -and $p.Out -match "Plan: [0-9a-f]{12}" -and -not (Item $ids.A).Complete -and -not (Item $ids.C).Complete)
    Check "nothing changed by the previews" ((Due $ids.A) -and (Due $ids.B) -and (Due $ids.C) -and -not (Item $ids.G).Complete)

    # Plan lock
    $r = Run "outlook-tasks-triage.ps1" @("-Search", $tag, "-Older", "30", "-Complete", "-Confirm")
    Check "-Confirm without -Plan is refused" ($r.Exit -eq 1 -and $r.Out -match "needs -Plan" -and -not (Item $ids.C).Complete)
    $stale = (RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, "-Older", "30", "-Complete")).Obj.plan
    $null = Run "outlook-tasks-update.ps1" @("-EntryID", $ids.A, "-DueDate", $today.AddDays(-50).ToString("yyyy-MM-dd"))   # A now matches -Older 30 too
    $r = Run "outlook-tasks-triage.ps1" @("-Search", $tag, "-Older", "30", "-Complete", "-Confirm", "-Plan", $stale)
    Check "a stale plan is refused and nothing is done" ($r.Exit -eq 1 -and $r.Out -match "does not match" -and -not (Item $ids.C).Complete -and -not (Item $ids.A).Complete)
    $null = Run "outlook-tasks-update.ps1" @("-EntryID", $ids.A, "-DueDate", $today.AddDays(-3).ToString("yyyy-MM-dd"))

    # Apply with the current plan: undo record
    $plan = (RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, "-Older", "30", "-Complete")).Obj.plan
    $r = RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, "-Older", "30", "-Complete", "-Confirm", "-Plan", $plan)
    $u = @($r.Obj.undo)
    Check "-Older 30 -Complete with the plan completes only the 40-day task" ($r.Exit -eq 0 -and $r.Obj.changed -eq 1 -and $r.Obj.failed -eq 0 -and (Item $ids.C).Complete -and -not (Item $ids.A).Complete -and -not (Item $ids.B).Complete)
    Check "undo record: EntryID, old due, old status" ($u.Count -eq 1 -and $u[0].EntryID -eq $ids.C -and $u[0].OldDue -eq $today.AddDays(-40).ToString("yyyy-MM-dd") -and $u[0].OldStatus -eq "NotStarted" -and @($r.Obj.failures).Count -eq 0)

    # -MaxTasks
    $pv = RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, "-ClearDue", "-MaxTasks", "1")
    $r = Run "outlook-tasks-triage.ps1" @("-Search", $tag, "-ClearDue", "-MaxTasks", "1", "-Confirm", "-Plan", $pv.Obj.plan)
    Check "more tasks than -MaxTasks is refused" ($pv.Obj.overMaxTasks -and $pv.Obj.wouldChange -eq 2 -and $r.Exit -eq 1 -and $r.Out -match "MaxTasks" -and (Due $ids.A) -and (Due $ids.B))

    # Reschedule, then clear
    $new = $today.AddDays(7)
    $plan = (RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, "-Category", "SuiteTriage", "-Reschedule", $new.ToString("yyyy-MM-dd"))).Obj.plan
    $r = RunJson "outlook-tasks-triage.ps1" @("-Search", $tag, "-Category", "SuiteTriage", "-Reschedule", $new.ToString("yyyy-MM-dd"), "-Confirm", "-Plan", $plan)
    Check "-Category -Reschedule with the plan moves only the open SuiteTriage task" ($r.Exit -eq 0 -and $r.Obj.changed -eq 1 -and (Due $ids.A) -eq $new -and (Due $ids.B) -eq $today.AddDays(-10))
    $t = Run "outlook-tasks-triage.ps1" @("-Search", $tag, "-ClearDue")
    $plan = [regex]::Match($t.Out, "Plan: ([0-9a-f]{12})").Groups[1].Value
    $r = Run "outlook-tasks-triage.ps1" @("-Search", $tag, "-ClearDue", "-Confirm", "-Plan", $plan)
    Check "-ClearDue (text) clears B, skips D and G, prints the undo record" ($r.Exit -eq 0 -and $r.Out -match "1 changed" -and $r.Out -match "1 recurring skipped" -and $r.Out -match "1 assigned skipped" -and $r.Out -match "Undo record" -and $r.Out -match [regex]::Escape($ids.B) -and -not (Due $ids.B) -and (Due $ids.D) -and (Due $ids.G) -and -not (Item $ids.G).Complete)
    $rep = RunJson "outlook-tasks-triage.ps1" @("-Search", $tag)
    Check "afterwards only the recurring and the assigned task are overdue" ($rep.Obj.matched -eq 2)
    Check "real overdue tasks untouched ($realBefore)" ((RealOverdue) -eq $realBefore) "now $(RealOverdue)"
}
finally {
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $savedGate
    # Delete, then remove the copies from Deleted Items by exact subject.
    $del = $Namespace.GetDefaultFolder(3)
    $subjects = New-Object System.Collections.Generic.List[string]
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
