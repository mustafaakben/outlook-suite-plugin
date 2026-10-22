param(
    [int]$Older = 0,
    [string]$Category = "",
    [string]$Search = "",
    [datetime]$Reschedule,
    [switch]$Complete,
    [switch]$ClearDue,
    [switch]$All,
    [switch]$IncludeParked,
    [string]$Plan = "",
    [int]$MaxTasks = 100,
    [int]$Limit = 5,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Tasks Triage: work through overdue tasks in bulk.
# Without an action it reports: overdue open tasks grouped by age (under 1 week, 1 to 4 weeks,
# over a month) and by category, with counts and a few examples per group.
# Waiting and Deferred tasks are left out unless -IncludeParked.
#
# Actions apply to the overdue tasks that match the filters:
#   -Reschedule <date>   new due date (start date moved back if it would be after it)
#   -Complete            mark complete
#   -ClearDue            remove the due date (and start date)
# Filters: -Older N (overdue by more than N days), -Category <name>, -Search <subject text>.
# An action needs at least one filter, or -All for every overdue task.
#
# Two steps, so what you approve is exactly what runs:
#   1. Run the action without -Confirm: a preview with counts and a plan code (a hash of the
#      action, the filters and the exact tasks).
#   2. Run it again with -Confirm -Plan <code>. If the tasks changed in between, the plan code no
#      longer matches and nothing is done.
# At most -MaxTasks tasks (default 100) per run; raise it explicitly for more.
# Never touched: recurring tasks (changing or completing one moves the series or spawns the next)
# and assigned tasks (Delegator set or completion reports to someone: completing them can email
# that person). Both are counted in the preview.
# The result lists every changed task with its EntryID and old due and start dates, so a run
# can be reversed with outlook-tasks-update.ps1; failures list the EntryID and the error, and
# the exit code is 1 when anything failed.
# Usage: .\outlook-tasks-triage.ps1
#        .\outlook-tasks-triage.ps1 -Older 30 -ClearDue                              (preview, prints the plan)
#        .\outlook-tasks-triage.ps1 -Older 30 -ClearDue -Confirm -Plan 3f9a1c2b7d10
#        .\outlook-tasks-triage.ps1 -Category "Teaching" -Reschedule 2026-11-06      (preview)

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")

$Buckets = @("Under 1 week", "1 to 4 weeks", "Over a month")
function Get-AgeBucket([int]$Days) {
    if ($Days -lt 7) { return $Buckets[0] }
    if ($Days -le 30) { return $Buckets[1] }
    return $Buckets[2]
}

# Plan code: the action, its date, the filters and the sorted EntryIDs, hashed (12 hex characters).
function Get-TriagePlan([string]$Action, $Targets, $Filters, [string]$When) {
    $ids = @($Targets | ForEach-Object { [string]$_.EntryID } | Sort-Object)
    $text = "$Action|$When|" + (($Filters.Keys | ForEach-Object { "$_=$($Filters[$_])" }) -join ";") + "|" + ($ids -join ",")
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($text)) } finally { $sha.Dispose() }
    return (ConvertTo-HexId $hash).Substring(0, 12).ToLower()
}

# The bulk gate. Its only output is the bool; the caller prints the preview.
function Confirm-TriageAction {
    param([string]$What, [int]$Count, [bool]$Confirmed)
    if ($Confirmed -or -not (Test-ConfirmRequired -Action "bulk")) { return $true }
    return $false
}

try {
    $actions = @($Complete, $ClearDue, [bool]$PSBoundParameters.ContainsKey("Reschedule")) | Where-Object { $_ }
    if ($actions.Count -gt 1) { Stop-WithError "Choose one action: -Reschedule, -Complete or -ClearDue." }
    $action = if ($Complete) { "complete" } elseif ($ClearDue) { "clear-due" } elseif ($PSBoundParameters.ContainsKey("Reschedule")) { "reschedule" } else { "" }
    $filtered = $Older -gt 0 -or $Category -or $Search
    if ($action -and -not $filtered -and -not $All) {
        Stop-WithError "A bulk action needs a filter (-Older, -Category, -Search) or -All for every overdue task."
    }
    if ($action -eq "reschedule" -and $Reschedule.Date -lt (Get-Date).Date) { Stop-WithError "-Reschedule is in the past; the tasks would stay overdue." }
    if ($MaxTasks -lt 1) { Stop-WithError "-MaxTasks must be at least 1." }

    $null = Connect-Outlook
    $today = (Get-Date).Date
    $overdue = @((Get-TaskRecords -Namespace $Namespace).Records | Where-Object { Test-TaskDue $_ "Overdue" })
    foreach ($r in $overdue) {
        $r.DaysOverdue = [int]($today - $r.Due.Date).TotalDays
        $r.Age = Get-AgeBucket $r.DaysOverdue
        $cats = @(([string]$r.Categories) -split '\s*[,;]\s*' | Where-Object { $_ })
        $r.CategoryList = $(if ($cats.Count) { $cats } else { @("(none)") })
    }
    $parked = @($overdue | Where-Object { $_.Status -in "Waiting", "Deferred" })
    $pool = if ($IncludeParked) { $overdue } else { @($overdue | Where-Object { $_.Status -notin "Waiting", "Deferred" }) }
    $selected = @($pool | Where-Object {
        $_.DaysOverdue -gt $Older -and
        (-not $Category -or $_.CategoryList -contains $Category) -and
        (-not $Search -or $_.Subject.IndexOf($Search, [StringComparison]::OrdinalIgnoreCase) -ge 0)
    } | Sort-Object { $_.DaysOverdue } -Descending)

    $byAge = [ordered]@{}; foreach ($b in $Buckets) { $byAge[$b] = @($selected | Where-Object { $_.Age -eq $b }) }
    $byCat = @{}
    foreach ($r in $selected) { foreach ($c in $r.CategoryList) { if (-not $byCat.ContainsKey($c)) { $byCat[$c] = 0 }; $byCat[$c]++ } }
    $catRows = @($byCat.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { [ordered]@{ Category = $_.Key; Count = $_.Value } })

    $filters = [ordered]@{ Older = $Older; Category = $Category; Search = $Search; All = [bool]$All; IncludeParked = [bool]$IncludeParked }
    $toJsonTask = { param($r) $o = ConvertTo-TaskJson $r; $o.DaysOverdue = $r.DaysOverdue; $o.Age = $r.Age; $o.Remove("CategoryList"); $o }
    $parkedNote = if ($IncludeParked) { "" } else { "$($parked.Count) Waiting/Deferred left out (-IncludeParked)" }

    if (-not $action) {
        if ($Json) {
            Write-Json ([ordered]@{ overdue = $overdue.Count; matched = $selected.Count; parkedExcluded = $(if ($IncludeParked) { 0 } else { $parked.Count }); filters = $filters
                byAge = [ordered]@{ "Under 1 week" = $byAge[$Buckets[0]].Count; "1 to 4 weeks" = $byAge[$Buckets[1]].Count; "Over a month" = $byAge[$Buckets[2]].Count }
                byCategory = $catRows
                examples = [ordered]@{
                    "Under 1 week" = @($byAge[$Buckets[0]] | Select-Object -First $Limit | ForEach-Object { & $toJsonTask $_ })
                    "1 to 4 weeks" = @($byAge[$Buckets[1]] | Select-Object -First $Limit | ForEach-Object { & $toJsonTask $_ })
                    "Over a month" = @($byAge[$Buckets[2]] | Select-Object -First $Limit | ForEach-Object { & $toJsonTask $_ })
                } })
            exit 0
        }
        Write-Host "=== OVERDUE TASKS: $($selected.Count)$(if ($filtered) { " matching (of $($overdue.Count))" }) ===" -ForegroundColor Cyan
        if ($parkedNote) { Write-Host "  ($parkedNote)" -ForegroundColor Gray }
        Write-Host "`n--- By age ---" -ForegroundColor Cyan
        foreach ($b in $Buckets) {
            $list = $byAge[$b]
            Write-Host ("  {0,-14} {1,5}" -f $b, $list.Count) -ForegroundColor Yellow
            foreach ($r in ($list | Select-Object -First $Limit)) {
                Write-Host ("      {0,4} days  {1}{2}" -f $r.DaysOverdue, $r.Subject, $(if ($r.Categories) { "  [$($r.Categories)]" } else { "" }))
                Write-Host "                 EntryID: $($r.EntryID)" -ForegroundColor DarkGray
            }
        }
        Write-Host "`n--- By category ---" -ForegroundColor Cyan
        foreach ($row in $catRows) { Write-Host ("  {0,-30} {1,5}" -f $row.Category, $row.Count) }
        Write-Host "`nNext: -Older 30 -ClearDue | -Category <name> -Reschedule <date> | -Older 90 -Complete (each previews first and prints a plan code)" -ForegroundColor DarkGray
        exit 0
    }

    # --- Bulk action: work out exactly which tasks it would change ---
    $recurring = @($selected | Where-Object { $_.Recurring })
    $candidates = @($selected | Where-Object { -not $_.Recurring })
    # Assigned tasks: completing or changing them can send a status report to the assigner.
    $delegated = New-Object System.Collections.Generic.List[object]
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($r in $candidates) {
        $isDelegated = $false
        try {
            $t = $Namespace.GetItemFromID($r.EntryID)
            $isDelegated = [bool]([string]$t.Delegator) -or [bool]([string]$t.StatusOnCompletionRecipients) -or [int]$t.DelegationState -ne 0
        } catch { $isDelegated = $true }   # cannot check it: leave it alone
        if ($isDelegated) { $delegated.Add($r) } else { $targets.Add($r) }
    }
    $targets = @($targets.ToArray())
    $when = if ($action -eq "reschedule") { $Reschedule.ToString("yyyy-MM-dd") } else { "" }
    $planCode = Get-TriagePlan -Action $action -Targets $targets -Filters $filters -When $when
    $label = switch ($action) {
        "reschedule" { "RESCHEDULE to $when" }
        "complete" { "MARK COMPLETE" }
        "clear-due" { "CLEAR DUE DATE" }
    }
    $overCap = $targets.Count -gt $MaxTasks
    $summary = [ordered]@{ action = $action; matched = $selected.Count; wouldChange = $targets.Count
        skippedRecurring = $recurring.Count; skippedAssigned = $delegated.Count
        parkedExcluded = $(if ($IncludeParked) { 0 } else { $parked.Count }); maxTasks = $MaxTasks; plan = $planCode; filters = $filters }

    # Fail-closed: anything but exactly $true from the gate keeps this a preview.
    $gateClosed = (Confirm-TriageAction -What $label -Count $targets.Count -Confirmed $Confirm) -ne $true
    $planOk = $Plan -and $Plan.Trim().ToLower() -eq $planCode
    if (-not $gateClosed -and $Confirm -and -not $Plan) { Stop-WithError "-Confirm needs -Plan <code> from the preview, so that exactly the previewed tasks are changed." -Hint "Run the same command without -Confirm to see the plan code ($($targets.Count) task(s) now)." }
    if (-not $gateClosed -and $Plan -and -not $planOk) { Stop-WithError "The plan code does not match: the matching tasks changed since the preview (now $($targets.Count) task(s), plan $planCode). Nothing was changed." -Hint "Preview again and check the list before confirming." }
    if (-not $gateClosed -and $planOk -and $overCap) { Stop-WithError "$($targets.Count) tasks exceed -MaxTasks $MaxTasks. Nothing was changed." -Hint "Narrow the filters, or pass -MaxTasks $($targets.Count) to allow it." }

    if ($gateClosed -or -not $planOk) {
        # Preview: counts, examples and the plan code. Nothing is changed.
        if ($Json) {
            $o = [ordered]@{}; foreach ($k in $summary.Keys) { $o[$k] = $summary[$k] }
            $o.result = "preview"; $o.done = $false; $o.overMaxTasks = $overCap
            $o.byAge = [ordered]@{ "Under 1 week" = @($targets | Where-Object { $_.Age -eq $Buckets[0] }).Count; "1 to 4 weeks" = @($targets | Where-Object { $_.Age -eq $Buckets[1] }).Count; "Over a month" = @($targets | Where-Object { $_.Age -eq $Buckets[2] }).Count }
            $o.examples = @($targets | Select-Object -First $Limit | ForEach-Object { & $toJsonTask $_ })
            $o.hint = "Nothing was changed. To apply exactly this plan: the same command with -Confirm -Plan $planCode$(if ($overCap) { " -MaxTasks $($targets.Count)" })"
            Write-Json $o
            exit 0
        }
        Write-Host "=== ${label}: $($targets.Count) task(s) (PREVIEW) ===" -ForegroundColor Yellow
        foreach ($r in ($targets | Select-Object -First $Limit)) { Write-Host ("      {0,4} days  {1}" -f $r.DaysOverdue, $r.Subject) }
        if ($targets.Count -gt $Limit) { Write-Host "      ... and $($targets.Count - $Limit) more" -ForegroundColor Gray }
        if ($recurring.Count) { Write-Host "  Recurring tasks skipped: $($recurring.Count)" -ForegroundColor Gray }
        if ($delegated.Count) { Write-Host "  Assigned tasks skipped: $($delegated.Count)" -ForegroundColor Gray }
        if ($parkedNote) { Write-Host "  $parkedNote" -ForegroundColor Gray }
        if ($overCap) { Write-Host "  More than -MaxTasks ${MaxTasks}: narrow the filters or add -MaxTasks $($targets.Count)." -ForegroundColor Yellow }
        Write-Host ""
        Write-Host "PREVIEW ONLY: nothing was changed. Plan: $planCode" -ForegroundColor Yellow
        Write-Host "To apply exactly this plan, run the same command with -Confirm -Plan $planCode$(if ($overCap) { " -MaxTasks $($targets.Count)" })" -ForegroundColor Gray
        exit 0
    }

    # --- Apply, keeping an undo record ---
    $done = New-Object System.Collections.Generic.List[object]
    $failures = New-Object System.Collections.Generic.List[object]
    foreach ($r in $targets) {
        $oldDue = $r.Due; $oldStart = $r.Start
        try {
            $t = $Namespace.GetItemFromID($r.EntryID)
            switch ($action) {
                "complete" { $t.MarkComplete() }
                "clear-due" {
                    $t.DueDate = [datetime]"4501-01-01"
                    $t.StartDate = [datetime]"4501-01-01"
                    $t.Save()
                }
                "reschedule" {
                    if (([datetime]$t.StartDate).Year -lt 4500 -and ([datetime]$t.StartDate).Date -gt $Reschedule.Date) { $t.StartDate = $Reschedule.Date }
                    $t.DueDate = $Reschedule.Date
                    $t.Save()
                }
            }
            $undo = [ordered]@{ EntryID = $r.EntryID; Subject = $r.Subject; OldDue = $(if ($oldDue) { Format-TaskDate $oldDue } else { $null }); OldStart = $(if ($oldStart) { Format-TaskDate $oldStart } else { $null }); OldStatus = $r.Status }
            $done.Add($undo)
        } catch {
            $fail = [ordered]@{ EntryID = $r.EntryID; Subject = $r.Subject; Error = $_.Exception.Message }
            $failures.Add($fail)
        }
    }

    if ($Json) {
        $o = [ordered]@{}; foreach ($k in $summary.Keys) { $o[$k] = $summary[$k] }
        $o.result = "done"; $o.done = $true; $o.changed = $done.Count; $o.failed = $failures.Count
        $o.undo = @($done.ToArray()); $o.failures = @($failures.ToArray())
        Write-Json $o
    } else {
        Write-Host "=== ${label}: $($done.Count) changed$(if ($failures.Count) { ", $($failures.Count) failed" })$(if ($recurring.Count) { ", $($recurring.Count) recurring skipped" })$(if ($delegated.Count) { ", $($delegated.Count) assigned skipped" }) ===" -ForegroundColor Green
        Write-Host "Plan: $planCode" -ForegroundColor Gray
        if ($done.Count) {
            Write-Host "`n--- Undo record (old dates; restore with outlook-tasks-update.ps1 -EntryID <id> -DueDate <old due> -StartDate <old start>, or -Status) ---" -ForegroundColor Cyan
            foreach ($u in $done) { Write-Host "  $($u.EntryID)  due $(if ($u.OldDue) { $u.OldDue } else { '(none)' })  start $(if ($u.OldStart) { $u.OldStart } else { '(none)' })  status $($u.OldStatus)  $($u.Subject)" }
        }
        if ($failures.Count) {
            Write-Host "`n--- Failed ---" -ForegroundColor Red
            foreach ($f in $failures) { Write-Host "  $($f.EntryID)  $($f.Error)" -ForegroundColor Red }
        }
    }
    if ($failures.Count -gt 0) { exit 1 }
    exit 0
}
catch {
    Stop-WithError $_
}
