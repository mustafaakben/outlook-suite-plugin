<#
.SYNOPSIS
  Live test for outlook-calendar-free.ps1 and outlook-calendar-week.ps1 (when present).

.DESCRIPTION
  1. Read-only cross-check: every offered slot avoids every blocking meeting in the real
     calendar, sits inside working hours, and carries the tier of its day.
  2. Tier rules on synthetic records, then live: a "[suite-test]" hold is counted and an
     all-day Out of Office day gets no slots.
  Prints PASS/FAIL and counts only. Deletes everything it creates.
#>
param(
    [ValidateSet("pwsh", "powershell")]
    [string]$Exe = "pwsh",
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-calendar/scripts")
)

$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
$null = Connect-Outlook
$prevConfirm = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"   # gate forced on; restored at the end

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $out = & $Exe -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out; Ms = $sw.ElapsedMilliseconds }
}
# Removes this run's items from the calendar and from Deleted Items (a deleted appointment
# lands there), then checks from a separate process, which does not see this session's cache.
function Remove-TestItems([string]$prefix) {
    $q = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%$($prefix.Replace("'", "''"))%'"   # also "FW: <prefix>..." forwards
    foreach ($fid in 9, 16, 3) {   # calendar and drafts first: deleted items move to Deleted Items
        foreach ($i in @($Namespace.GetDefaultFolder($fid).Items.Restrict($q))) { try { $i.Delete() } catch { } }
    }
    $env:SUITE_TEST_QUERY = $q
    $check = '$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI"); $n = 0; foreach ($fid in 9, 16, 3, 4, 5) { $n += $ns.GetDefaultFolder($fid).Items.Restrict($env:SUITE_TEST_QUERY).Count }; $n'
    return [int](& pwsh -NoProfile -Command $check)
}
# Items of this run in Outbox or Sent, counted by a separate process: must stay 0 (nothing is
# ever sent by these tests; a preview that really sends is the worst bug this suite can have).
function Get-SentTestItems([string]$prefix) {
    $env:SUITE_TEST_QUERY = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%$($prefix.Replace("'", "''"))%'"
    return [int](& pwsh -NoProfile -Command '$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI"); $ns.GetDefaultFolder(4).Items.Restrict($env:SUITE_TEST_QUERY).Count + $ns.GetDefaultFolder(5).Items.Restrict($env:SUITE_TEST_QUERY).Count')
}
function J($r) { try { return ($r.Out | ConvertFrom-Json) } catch { return $null } }
function D($iso) { return ([datetimeoffset]::Parse($iso)).LocalDateTime }
function Find-Test([datetime]$from, [datetime]$to) {
    $items = $Namespace.GetDefaultFolder(9).Items
    $items.IncludeRecurrences = $true
    $f = "[Start] < '$($to.ToString('MM/dd/yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture))' AND [End] > '$($from.ToString('MM/dd/yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture))'"
    return @(foreach ($e in $items.Restrict($f)) { if ($e.Subject -like "``[suite-test``]*") { $e } })
}

Write-Host "Running under $Exe" -ForegroundColor Gray

# 1. Cross-check against the real calendar (read only)
$r = Run "outlook-calendar-free.ps1" @("-Days", "14", "-Duration", "30", "-PerDay", "3", "-Top", "50", "-Json"); $free = J $r
Check "free -Days 14 -Json" ($r.Exit -eq 0 -and $free) "exit=$($r.Exit)"
Write-Host "      free 14d: $(@($free.Slots).Count) slots, $($r.Ms) ms" -ForegroundColor Gray
$r = Run "outlook-calendar-list.ps1" @("-From", (Get-Date).ToString("yyyy-MM-dd"), "-Days", "15", "-Limit", "2000", "-Json"); $list = J $r
$meetings = @($list.Events | Where-Object { $_.Blocking -and $_.Kind -eq "meeting" })
$clash = 0; $outside = 0; $past = 0
foreach ($s in $free.Slots) {
    $a = D $s.Start; $b = D $s.End
    if ($a -lt (Get-Date)) { $past++ }
    if ($a.TimeOfDay -lt [TimeSpan]"09:00" -or $b.TimeOfDay -gt [TimeSpan]"17:00" -or $a.Date -ne $b.Date) { $outside++ }
    foreach ($m in $meetings) { if ((D $m.Start) -lt $b -and $a -lt (D $m.End)) { $clash++ } }
}
Check "no slot overlaps a blocking meeting" ($clash -eq 0) "clashes=$clash"
Check "every slot is inside 9 to 5 and in the future" ($outside -eq 0 -and $past -eq 0) "outside=$outside past=$past"
$byDate = @{}; foreach ($d in $free.Week) { $byDate[(D $d.Date).ToString("yyyy-MM-dd")] = $d }
$wrongTier = @($free.Slots | Where-Object { $byDate[(D $_.Start).ToString("yyyy-MM-dd")].Tier -ne $_.Tier }).Count
Check "every slot carries its day's tier" ($wrongTier -eq 0)
$order = @($free.Slots | ForEach-Object { $_.Tier })
Check "slots are ordered best tier first" ((($order | Sort-Object) -join ",") -eq ($order -join ","))
if (-not $free.BusyDaysOnly) { Check "no busy-day slot while better days have room" (@($free.Slots | Where-Object { $_.Tier -eq 4 }).Count -eq 0) }

# With the user as an attendee, their own free/busy is used too: slots must also avoid busy holds.
$me = $Namespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
$r = Run "outlook-calendar-free.ps1" @("-Days", "7", "-Attendees", $me, "-Json"); $j = J $r
$busyAll = @($list.Events | Where-Object { $_.Busy -in @("Busy", "Out of Office") -and -not $_.AllDay })
$clash = 0
foreach ($s in $j.Slots) { $a = D $s.Start; $b = D $s.End; foreach ($m in $busyAll) { if ((D $m.Start) -lt $b -and $a -lt (D $m.End)) { $clash++ } } }
Check "attendee free/busy is honored (self as attendee avoids busy holds)" ($r.Exit -eq 0 -and $clash -eq 0) "clashes=$clash"
$r = Run "outlook-calendar-free.ps1" @("-Attendees", "zz-no-such-person-zz", "-Days", "2")
Check "unknown attendee is reported, search still runs" ($r.Exit -eq 0 -and $r.Out -match "zz-no-such-person-zz")
$r = Run "outlook-calendar-free.ps1" @("-Days", "3", "-Duration", "600")
Check "impossible duration: widens, then lists the skipped days and options, exit 0" ($r.Exit -eq 0 -and $r.Out -match "is taken" -and $r.Out -match "Options:" -and $r.Out -match "widened")

# 1a. Steering flags (3.1): tier filter, busy days, spread, time-of-day preference, skipped days.
$tierPresent = @($free.Slots | ForEach-Object { [int]$_.Tier } | Sort-Object -Unique)
if ($tierPresent.Count) {
    $pick = $tierPresent[-1]
    $r = Run "outlook-calendar-free.ps1" @("-Days", "14", "-Tier", "$pick", "-Json"); $j = J $r
    Check "-Tier $pick offers only that tier" ($r.Exit -eq 0 -and $j -and @($j.Slots | Where-Object { $_.Tier -ne $pick }).Count -eq 0 -and @($j.Tiers) -join "," -eq "$pick")
}
$r = Run "outlook-calendar-free.ps1" @("-Days", "14", "-Tier", "1,3", "-Json"); $j = J $r
Check "-Tier '1,3' (one string) is split" ($r.Exit -eq 0 -and $j -and (@($j.Tiers) -join ",") -eq "1,3" -and @($j.Slots | Where-Object { $_.Tier -notin 1, 3 }).Count -eq 0)
$r = Run "outlook-calendar-free.ps1" @("-Days", "14", "-IncludeBusy", "-Top", "50", "-Json"); $j = J $r
Check "-IncludeBusy keeps every normal slot" ($r.Exit -eq 0 -and $j -and @($j.Slots).Count -ge @($free.Slots | Where-Object { $_.Tier -le 3 }).Count -and (@($j.Tiers) -join ",") -eq "1,2,3,4")
$r = Run "outlook-calendar-free.ps1" @("-Days", "14", "-Propose", "-Spread", "-Count", "4", "-Json"); $j = J $r
$tiersOffered = @($free.Slots | Where-Object { $_.Tier -le 3 } | ForEach-Object { $_.Tier } | Sort-Object -Unique).Count
Check "-Spread -Propose covers every offered tier first" ($r.Exit -eq 0 -and $j -and @($j.Proposals | ForEach-Object { $_.Tier } | Sort-Object -Unique).Count -ge [Math]::Min($tiersOffered, 4))
$r = Run "outlook-calendar-free.ps1" @("-Days", "14", "-Prefer", "Afternoon", "-Json"); $j = J $r
$pmA = @($j.Slots | Where-Object { (D $_.Start).Hour -ge 12 }).Count; $pm0 = @($free.Slots | Where-Object { (D $_.Start).Hour -ge 12 }).Count
Check "-Prefer Afternoon never offers fewer afternoon slots" ($r.Exit -eq 0 -and $j -and $pmA -ge [Math]::Min($pm0, @($j.Slots).Count)) "pm=$pmA vs $pm0"
$r = Run "outlook-calendar-free.ps1" @("-Days", "14", "-SkipDays", "Mon,Tue", "-Json"); $j = J $r
Check "-SkipDays Mon,Tue leaves those days out" ($r.Exit -eq 0 -and $j -and @($j.Slots | Where-Object { (D $_.Start).DayOfWeek -in "Monday", "Tuesday" }).Count -eq 0)
$r = Run "outlook-calendar-free.ps1" @("-Tier", "5", "-Days", "2")
Check "-Tier 5 is rejected" ($r.Exit -ne 0 -and $r.Out -match "Tier")
$r = Run "outlook-calendar-free.ps1" @("-SkipDays", "Xyz", "-Days", "2")
Check "-SkipDays with a bad value is rejected" ($r.Exit -ne 0 -and $r.Out -match "SkipDays")

# 1b. Independent Kind check: every meeting the user organizes in the next 30 days, classified
#     from the table path (list) and from the opened item, must agree.
. (Join-Path $ScriptsDir "_calendar.ps1")
$sw = [Diagnostics.Stopwatch]::StartNew()
$ev = Get-CalendarEvents -Namespace $Namespace -From (Get-Date).Date -To (Get-Date).Date.AddDays(30)
$ms1 = 0; $mismatch = 0; $seenIds = @{}
foreach ($e in $ev) {
    if (-not $e.EntryID -or $seenIds.ContainsKey($e.EntryID)) { continue }
    $it = $Namespace.GetItemFromID($e.EntryID)
    if ([int]$it.MeetingStatus -ne 1) { continue }
    $seenIds[$e.EntryID] = 1; $ms1++
    if ((ConvertTo-EventRecord $it).Kind -ne $e.Kind) { $mismatch++ }
}
Write-Host "      organizer events: $ms1, Kind mismatches table vs item: $mismatch ($($sw.ElapsedMilliseconds) ms)" -ForegroundColor Gray
Check "Kind agrees between the table path and the opened item (organizer meetings)" ($mismatch -eq 0) "mismatches=$mismatch"
Check "1:1 rule: a meeting with one other invitee is a meeting, not a hold" ((Get-OtherAttendeeCount -Names @("Jane Doe") -Organizer "Test Organizer") -eq 1 -and (Get-EventKind -AllDay $false -MeetingStatus 1 -Attendees 1 -Subject "1:1" -Busy "Busy" -Private $false -FirmPatterns @()) -eq "meeting")

# 1c. Propose: email-ready lines, HTML block, display time zone
$r = Run "outlook-calendar-free.ps1" @("-Propose", "-Count", "3")
$lines = @($r.Out -split "`r?`n" | Where-Object { $_ -match '\S' -and $_ -notmatch '^NOT CHECKED' })
Check "-Propose prints up to 3 lines like 'Mon, Oct 26, 10:30 to 11:30 a.m.'" ($r.Exit -eq 0 -and $lines.Count -ge 1 -and $lines.Count -le 3 -and @($lines | Where-Object { $_ -notmatch '^[A-Z][a-z]{2}, [A-Z][a-z]{2} \d{1,2}, \d{1,2}:\d{2}( [ap]\.m\.)? to \d{1,2}:\d{2} [ap]\.m\.$' }).Count -eq 0) "lines=$($lines.Count)"
Check "-Propose never uses en or em dashes" ($r.Out -notmatch "[\u2013\u2014]")
$r = Run "outlook-calendar-free.ps1" @("-Propose", "-Html")
Check "-Propose -Html is one Outlook-native <div><ul> block" ($r.Exit -eq 0 -and $r.Out.Trim() -match '^<div style="font-family: Aptos[^"]*rgb\(0, 0, 0\);" class="elementToProof"><ul>(<li [^>]*class="elementToProof">[^<]+</li>)+</ul></div>$')
$r = Run "outlook-calendar-free.ps1" @("-Propose", "-Json", "-TimeZone", "Pacific"); $j = J $r
$p0 = @($j.Proposals)[0]
# Raw JSON text: PowerShell 7's ConvertFrom-Json turns ISO strings into local DateTimes.
$m = [regex]::Match($r.Out, '"Proposals":\[\{"Start":"[^"]*T(\d\d):(\d\d)[^"]*".*?"StartShown":"[^"]*T(\d\d):(\d\d):\d\d(-0[78]:00)"')
$shift = if ($m.Success) { [int]$m.Groups[1].Value - [int]$m.Groups[3].Value } else { -1 }
Check "-TimeZone Pacific shows times 3 hours earlier with a Pacific offset, JSON keeps both" ($r.Exit -eq 0 -and $p0 -and $p0.Text -like "*Pacific time" -and $shift -eq 3 -and $j.Html) "shift=$shift"
$r = Run "outlook-calendar-free.ps1" @("-From", (Get-Date).AddDays(-3).ToString("yyyy-MM-dd"), "-Days", "6", "-WorkStart", "9", "-WorkEnd", "4:00 PM", "-Json"); $j = J $r
Check "past -From days are skipped, '9' and '4:00 PM' accepted" ($r.Exit -eq 0 -and $j.WorkHours -eq "09:00-16:00" -and @($j.Notes).Count -ge 1 -and @($j.Slots | Where-Object { (D $_.Start) -lt (Get-Date) }).Count -eq 0)
$r = Run "outlook-calendar-free.ps1" @("-WorkStart", "nine-ish")
Check "unreadable -WorkStart is a clear error" ($r.Exit -eq 1 -and $r.Out -match "Cannot read -WorkStart")

# 1d. Attendee free/busy across the DST change: indexing by elapsed (UTC) time must give
#     the same answer as a direct per-day call (read-only, the user's own free/busy).
$dst = $null
for ($i = 1; $i -le 27; $i++) { $x = (Get-Date).Date.AddDays($i); if ($x.IsDaylightSavingTime() -ne (Get-Date).Date.IsDaylightSavingTime()) { $dst = $x; break } }
if ($dst) {
    $rcp = $Namespace.CreateRecipient($Namespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress); [void]$rcp.Resolve()
    $base = (Get-Date).Date; $all = [string]$rcp.FreeBusy($base, 30, $true)
    $day = $dst.AddDays(1); $direct = ([string]$rcp.FreeBusy($day, 30, $true)).Substring(0, 48)
    $utcIdx = [int][Math]::Floor(($day.ToUniversalTime() - $base.ToUniversalTime()).TotalMinutes / 30)
    Check "free/busy: UTC-elapsed index matches a direct call after DST ($($day.ToString('MMM d')))" ($all.Substring($utcIdx, 48) -eq $direct)
} else { Write-Host "SKIP  no DST change in the next 4 weeks" -ForegroundColor Yellow }
$r = Run "outlook-calendar-free.ps1" @("-Attendees", "zz-no-such-person-zz", "-Days", "2", "-Json"); $j = J $r
Check "unresolved attendee: NOT CHECKED in the header and on every slot" ($r.Exit -eq 0 -and @($j.NotChecked).Count -eq 1 -and @($j.Slots | Where-Object { @($_.NotChecked).Count -eq 0 }).Count -eq 0)

# 2a. Tier rules on synthetic records (deterministic)
function Ev($kind, [int]$h, [int]$len = 60, $busy = "Busy", [switch]$Cancelled, $response = "Accepted") {
    $s = ([datetime]"2030-01-07").AddHours($h)
    return [ordered]@{ Subject = "x"; Start = $s; End = $s.AddMinutes($len); Kind = $kind; AllDay = ($kind -eq "allday"); Busy = $busy; Response = $response
        Cancelled = [bool]$Cancelled; Blocking = ($busy -ne "Free" -and -not $Cancelled); ConflictsWith = @() }
}
$day = [datetime]"2030-01-07"
Check "tier: empty day is Open" ((Get-CalendarDay -Records @() -Date $day).Tier -eq 1)
Check "tier: only holds is Holds-only" ((Get-CalendarDay -Records @((Ev hold 9), (Ev hold 13)) -Date $day).Tier -eq 2)
Check "tier: one meeting is One-meeting" ((Get-CalendarDay -Records @((Ev meeting 9), (Ev hold 13)) -Date $day).Tier -eq 3)
Check "tier: two meetings is Busy (cap 2)" ((Get-CalendarDay -Records @((Ev meeting 9), (Ev meeting 13)) -Date $day).Tier -eq 4)
Check "tier: -MaxPerDay 3 keeps two meetings in tier 3" ((Get-CalendarDay -Records @((Ev meeting 9), (Ev meeting 13)) -Date $day -MaxPerDay 3).Tier -eq 3)
Check "tier: cancelled, declined and Free meetings do not count" ((Get-CalendarDay -Records @((Ev meeting 9 -Cancelled), (Ev meeting 11 -response "Declined"), (Ev meeting 13 -busy "Free")) -Date $day).Tier -eq 1)
Check "away: all-day Out of Office" ((Get-CalendarDay -Records @((Ev allday 0 1440 "Out of Office")) -Date $day).Away)
Check "away: all-day Free is not away" (-not (Get-CalendarDay -Records @((Ev allday 0 1440 "Free")) -Date $day).Away)
$fp = Get-FirmHoldPatterns -Override $null
function K($subject, $busy = "Busy", [bool]$private = $false, [int]$ms = 0, [int]$att = 0) { Get-EventKind -AllDay $false -MeetingStatus $ms -Attendees $att -Subject $subject -Busy $busy -Private $private -FirmPatterns $fp }
Check "firm: course code first (XYZ 101, ABC1234)" ((K "XYZ 101 Principles") -eq "firm" -and (K "ABC1234 lab") -eq "firm")
Check "firm: course code anywhere (ABC-1234-A)" ((K "Lecture ABC-1234-A Main Hall 120") -eq "firm")
Check "firm: class, office hours, exam, travel, flight" ((@("Strategy Class", "Office Hours", "Final exam", "Travel to Seattle", "Flight AA123") | Where-Object { (K $_) -ne "firm" }).Count -eq 0)
Check "firm: Out of Office time and private events" ((K "Dentist" "Out of Office") -eq "firm" -and (K "Personal" "Busy" $true) -eq "firm")
Check "hold: Prep block, Focus block, Planning, Lunch, TODO" ((@("Prep block", "Focus block", "Planning", "Lunch", "TODO: review paper") | Where-Object { (K $_) -ne "hold" }).Count -eq 0)
function KL($subject, $loc) { Get-EventKind -AllDay $false -MeetingStatus 0 -Attendees 0 -Subject $subject -Busy "Busy" -Private $false -FirmPatterns $fp -Location $loc }
Check "firm: personal appointments (dentist, doctor, Dr ..., haircut)" ((@("Dentist appointment", "Doctor visit", "Dr Appointment", "Haircut") | Where-Object { (K $_) -ne "firm" }).Count -eq 0)
Check "firm: a hold at a street address" ((KL "Pick up" "123 Main St, Burlington, NC 27215") -eq "firm" -and (KL "Lunch" "2 Elm Avenue") -eq "firm")
Check "hold: online or room locations stay holds" ((KL "Focus" "https://zoom.us/j/123") -eq "hold" -and (KL "Planning" "Microsoft Teams Meeting") -eq "hold" -and (KL "Prep block" "Main Hall 120") -eq "hold" -and (KL "Drafting paper" "") -eq "hold")
Check "meeting stays meeting even with a course code" ((K "ABC-1234 sync" "Busy" $false 3 4) -eq "meeting")
Check "-FirmHolds override replaces the patterns" ((Get-EventKind -AllDay $false -MeetingStatus 0 -Attendees 0 -Subject "Prep block" -Busy "Busy" -Private $false -FirmPatterns @("Prep")) -eq "firm")
$recs = @((Ev meeting 9 90), (Ev meeting 10), (Ev hold 9), (Ev meeting 12))
Set-CalendarConflicts $recs
Check "conflicts: overlapping meetings flagged, holds and later meetings not" ($recs[0].ConflictsWith.Count -eq 1 -and $recs[1].ConflictsWith.Count -eq 1 -and $recs[2].ConflictsWith.Count -eq 0 -and $recs[3].ConflictsWith.Count -eq 0)

# 2b. Live: a test hold and an all-day Out of Office on two weekdays about 90 days out
$mon = (Get-Date).Date.AddDays(90)
while ($mon.DayOfWeek -ne "Monday") { $mon = $mon.AddDays(1) }
$tag = "[suite-test] free $(Get-Date -Format 'MMddHHmmss')"
$range = @("-From", $mon.ToString("yyyy-MM-dd"), "-To", $mon.AddDays(1).ToString("yyyy-MM-dd"), "-Json")
try {
    $j = J (Run "outlook-calendar-free.ps1" $range)
    $before = @($j.Week | Where-Object { (D $_.Date) -eq $mon })[0]
    $null = Run "outlook-calendar-create.ps1" @("-Subject", "$tag hold", "-Start", $mon.AddHours(7).ToString("yyyy-MM-dd HH:mm"), "-Duration", "60", "-NoReminder")
    $null = Run "outlook-calendar-create.ps1" @("-Subject", "$tag away", "-Start", $mon.AddDays(1).ToString("yyyy-MM-dd"), "-AllDay", "-BusyStatus", "OutOfOffice", "-NoReminder")
    $firm = J (Run "outlook-calendar-create.ps1" @("-Subject", "$tag ZZZ-9999-A Principles", "-Start", $mon.AddHours(9).ToString("yyyy-MM-dd HH:mm"), "-End", $mon.AddHours(17).ToString("yyyy-MM-dd HH:mm"), "-NoReminder", "-Json"))
    Check "a course-code hold is created as Kind=firm" ($firm.Kind -eq "firm")
    $j = J (Run "outlook-calendar-free.ps1" $range)
    $after = @($j.Week | Where-Object { (D $_.Date) -eq $mon })[0]
    $tue = @($j.Week | Where-Object { (D $_.Date) -eq $mon.AddDays(1) })[0]
    Check "a new hold and a firm event are counted on their day" ($after.Holds -eq $before.Holds + 1 -and $after.Firm -eq $before.Firm + 1 -and $after.Meetings -eq $before.Meetings)
    Check "no slot is offered over the firm event (9 to 5 is fully taken)" (@($j.Slots | Where-Object { (D $_.Start).Date -eq $mon }).Count -eq 0)
    $l = J (Run "outlook-calendar-list.ps1" @("-From", $mon.ToString("yyyy-MM-dd"), "-To", $mon.ToString("yyyy-MM-dd"), "-Kind", "firm", "-Json"))
    Check "list -Kind firm shows it" (@($l.Events | Where-Object { $_.EntryID -eq $firm.EntryID }).Count -eq 1)
    $o = J (Run "outlook-calendar-free.ps1" ($range + @("-FirmHolds", "NO-SUCH-PATTERN")))
    Check "with -FirmHolds overridden it is a hold again and slots may cover it" (@($o.Slots | Where-Object { (D $_.Start).Date -eq $mon -and @($_.OverHold) -contains "$tag ZZZ-9999-A Principles" }).Count -ge 1)
    Check "all-day Out of Office day is Away with no slots" ($tue.Away -and @($j.Slots | Where-Object { (D $_.Start).Date -eq $mon.AddDays(1) }).Count -eq 0)
    if (Test-Path (Join-Path $ScriptsDir "outlook-calendar-week.ps1")) {
        $w = J (Run "outlook-calendar-week.ps1" @("-From", $mon.ToString("yyyy-MM-dd"), "-Json"))
        $wm = @($w.Days | Where-Object { (D $_.Date) -eq $mon })[0]
        $wt = @($w.Days | Where-Object { (D $_.Date) -eq $mon.AddDays(1) })[0]
        Check "week overview matches the free-time summary" ($wm.Holds -eq $after.Holds -and $wm.Meetings -eq $after.Meetings -and $wm.Tier -eq $after.Tier -and $wt.Away)
    }
}
finally {
    $sent = Get-SentTestItems $tag
    Check "nothing was sent: 0 test items in Outbox or Sent (checked from a new process)" ($sent -eq 0) "sent=$sent"
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $prevConfirm
    $left = Remove-TestItems $tag
    Check "cleanup: no test items left in Calendar or Deleted Items (checked from a new process)" ($left -eq 0) "left=$left"
}

Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green } else { Write-Host "$($script:fails) FAILED" -ForegroundColor Red; exit 1 }
