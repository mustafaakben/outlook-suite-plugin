<#
.SYNOPSIS
  Live test for outlook-agenda.ps1 (read-only brief). Prints PASS/FAIL, counts and timings only.

.DESCRIPTION
  Adds three disposable "[suite-test]" items (a 15-minute hold late today shown as Free with no
  attendees and no reminder, a task due today, an overdue task), checks that the agenda counts
  them in the right sections, cross-checks every count against an independent query, checks
  speed (under 2 s), then removes the test items including their copies in Deleted Items.
  Run under both shells:  pwsh -File Test-Agenda.ps1   /   powershell -File Test-Agenda.ps1 -Shell powershell
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
$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $ErrorActionPreference = "Continue"
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $out = & $Shell -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out; Seconds = $sw.Elapsed.TotalSeconds }
}
function Agenda([object[]]$a) {
    $r = Run "outlook-agenda.ps1" ($a + "-Json")
    $obj = $null; try { $obj = $r.Out | ConvertFrom-Json } catch { }
    return [pscustomobject]@{ Exit = $r.Exit; Obj = $obj; Seconds = $r.Seconds; Out = $r.Out }
}

$stamp = Get-Date -Format "HHmmss"
$today = (Get-Date).Date
$subjects = New-Object System.Collections.Generic.List[string]

try {
    $base = Agenda @()
    Check "agenda -Json runs (today)" ($base.Exit -eq 0 -and $base.Obj -and @($base.Obj.days).Count -eq 1 -and $base.Obj.from -eq $today.ToString("yyyy-MM-dd")) "exit=$($base.Exit)"

    # Test items.
    $hold = $Outlook.CreateItem(1)   # olAppointmentItem: no attendees = a hold
    $hold.Subject = "[suite-test] agenda hold $stamp"
    $hold.Start = $today.AddHours(23).AddMinutes(30); $hold.Duration = 15
    $hold.BusyStatus = 0; $hold.ReminderSet = $false
    $hold.Save(); $subjects.Add($hold.Subject)
    foreach ($t in @(@("[suite-test] agenda due $stamp", $today), @("[suite-test] agenda overdue $stamp", $today.AddDays(-2)))) {
        $task = $Outlook.CreateItem(3); $task.Subject = $t[0]; $task.DueDate = $t[1]; $task.Save(); $subjects.Add($t[0])
    }

    $a = Agenda @()
    $b = $base.Obj.counts; $c = $a.Obj.counts
    Check "hold counted as a hold" ($c.Holds -eq $b.Holds + 1 -and $c.Meetings -eq $b.Meetings) "holds $($b.Holds)->$($c.Holds)"
    Check "task due today counted" ($c.TasksDue -eq $b.TasksDue + 1 -and @($a.Obj.tasks.due | Where-Object { $_.Subject -eq "[suite-test] agenda due $stamp" }).Count -eq 1)
    Check "overdue task counted" ($c.TasksOverdue -eq $b.TasksOverdue + 1)
    $h = @($a.Obj.days[0].Holds | Where-Object { $_.Subject -eq "[suite-test] agenda hold $stamp" })
    Check "hold listed with kind and Free" ($h.Count -eq 1 -and $h[0].Kind -eq "hold" -and $h[0].Busy -eq "Free")

    # Independent cross-checks.
    $me = $Namespace.CurrentUser.Name
    $items = $Namespace.GetDefaultFolder(9).Items; $items.Sort("[Start]"); $items.IncludeRecurrences = $true
    $ci = [Globalization.CultureInfo]::InvariantCulture
    $found = $items.Restrict("[Start] < '$($today.AddDays(1).ToString('MM/dd/yyyy HH:mm', $ci))' AND [End] > '$($today.ToString('MM/dd/yyyy HH:mm', $ci))'")
    # Classification comes from the calendar skill; this independent walk checks its totals:
    # meetings = others invited, not Free, not cancelled or declined; every other timed event is a
    # hold or a firm event (classes, office hours, ...); all-day events apart.
    $mtg = 0; $own = 0; $all = 0
    $it = $found.GetFirst()
    while ($it) {
        if ($it.Start -ge $today.AddDays(1)) { break }
        $others = 0; foreach ($rc in $it.Recipients) { if ($rc.Name -ne $me) { $others++ } }
        $off = [int]$it.MeetingStatus -in 5, 7 -or [int]$it.ResponseStatus -eq 4
        if (-not $off) {
            if ($it.AllDayEvent) { $all++ }
            elseif ($others -gt 0) { if ([int]$it.BusyStatus -ne 0) { $mtg++ } }
            else { $own++ }
        }
        $it = $found.GetNext()
    }
    $d0 = $a.Obj.days[0]
    Check "meetings match an attendee-based count ($mtg)" ($c.Meetings -eq $mtg) "agenda=$($c.Meetings)"
    Check "holds + firm listed = no-attendee events ($own)" ((@($d0.Holds).Count + @($d0.Firm).Count) -eq $own) "holds=$(@($d0.Holds).Count) firm=$(@($d0.Firm).Count)"
    Check "all-day count ($all)" ($c.AllDay -eq $all)
    Check "day has a tier and per-day counts" ($d0.Tier -and $null -ne $d0.Counts.Meetings -and $null -ne $d0.Counts.Firm)
    $l = $null
    foreach ($try in 1..2) { $lr = Run "outlook-tasks-list.ps1" @("-Due", "Overdue", "-Json"); try { $l = $lr.Out | ConvertFrom-Json; break } catch { Start-Sleep -Seconds 1 } }
    Check "overdue tasks match tasks-list ($($l.total))" ($l -and $c.TasksOverdue -eq $l.total) "exit=$($lr.Exit)"
    $lf = (Run "outlook-tasks-list.ps1" @("-Flagged", "-Due", "Overdue", "-Limit", "5000", "-Json")).Out | ConvertFrom-Json
    $mailOver = @($lf.items | Where-Object { $_.Type -eq "Mail" }).Count
    Check "overdue flagged email matches tasks-list -Flagged ($mailOver)" ($c.FlaggedMailOverdue -eq $mailOver)
    $since = (Get-Date).AddDays(-3)
    $unread = $Namespace.GetDefaultFolder(6).Items.Restrict("[UnRead] = True AND [ReceivedTime] >= '$(Format-JetDate $since)'")
    $pri = 0
    foreach ($m in $unread) {
        $score = $null; try { $score = $m.PropertyAccessor.GetProperty($script:CopilotProp.Score) } catch { }
        if ("$score" -eq "5" -or [int]$m.Importance -eq 2) { $pri++ }
    }
    Check "priority unread matches a per-item count ($pri)" ($c.PriorityUnread -eq $pri) "agenda=$($c.PriorityUnread)"

    # Week, day parsing, text output, speed.
    $w = Agenda @("-Week")
    Check "agenda -Week: 7 days, today's numbers included" ($w.Exit -eq 0 -and @($w.Obj.days).Count -eq 7 -and $w.Obj.counts.Holds -ge $c.Holds -and $w.Obj.to -eq $today.AddDays(6).ToString("yyyy-MM-dd"))
    $wd = $today.AddDays(3)
    $n = Agenda @("-Day", $wd.DayOfWeek.ToString())
    Check "agenda -Day <weekday name>" ($n.Exit -eq 0 -and $n.Obj.from -eq $wd.ToString("yyyy-MM-dd"))
    $n = Agenda @("-Day", "tomorrow")
    Check "agenda -Day tomorrow" ($n.Exit -eq 0 -and $n.Obj.from -eq $today.AddDays(1).ToString("yyyy-MM-dd"))
    $bad = Run "outlook-agenda.ps1" @("-Day", "not-a-day")
    Check "agenda -Day with a bad value fails" ($bad.Exit -eq 1)
    $t = Run "outlook-agenda.ps1" @()
    Check "text output has every section" ($t.Exit -eq 0 -and $t.Out -match "=== AGENDA" -and $t.Out -match "--- Calendar ---" -and $t.Out -match "--- Tasks ---" -and $t.Out -match "--- Flagged email ---" -and $t.Out -match "--- Priority unread mail")
    $times = @(foreach ($i in 1..5) { (Run "outlook-agenda.ps1" @("-Week", "-Json")).Seconds })
    $median = ($times | Sort-Object)[2]   # median of 5: other Outlook clients can make single runs slow
    Check ("agenda -Week under 2 s (median {0:N2} s)" -f $median) ($median -lt 2)
}
finally {
    # Delete the test items, then remove their copies from Deleted Items by exact subject.
    foreach ($s in $subjects) {
        foreach ($fid in 9, 13) {
            foreach ($x in @($Namespace.GetDefaultFolder($fid).Items.Restrict("[Subject] = '$($s.Replace("'", "''"))'"))) { try { $x.Delete() } catch { } }
        }
    }
    $del = $Namespace.GetDefaultFolder(3)
    $left = New-Object System.Collections.Generic.List[string]; foreach ($s in $subjects) { $left.Add($s) }
    $n = 0
    for ($try = 0; $try -lt 10 -and $left.Count -gt 0; $try++) {
        foreach ($s in @($left)) {
            $hits = @($del.Items.Restrict("[Subject] = '$($s.Replace("'", "''"))'"))
            foreach ($x in $hits) { try { $x.Delete(); $n++ } catch { } }
            if ($hits.Count -gt 0) { [void]$left.Remove($s) }
        }
        if ($left.Count -gt 0) { Start-Sleep -Seconds 1 }
    }
    Write-Host "Cleanup: $n test item(s) removed$(if ($left.Count) { "; not found in Deleted Items: $($left.Count)" })." -ForegroundColor Gray
}
$sentAfter = $Namespace.GetDefaultFolder(5).Items.Count; $outboxAfter = $Namespace.GetDefaultFolder(4).Items.Count
Check "nothing sent (Sent Items $sentBefore -> $sentAfter, Outbox $outboxBefore -> $outboxAfter)" ($sentAfter -le $sentBefore -and $outboxAfter -le $outboxBefore)
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASS ($Shell)" -ForegroundColor Green } else { Write-Host "$($script:fails) FAILED ($Shell)" -ForegroundColor Red }
exit $script:fails
