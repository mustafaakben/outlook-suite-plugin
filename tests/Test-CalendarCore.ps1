<#
.SYNOPSIS
  Live test for the calendar scripts, on disposable "[suite-test]" holds only.

.DESCRIPTION
  Creates holds with no attendees about 60 days ahead at 6 AM, runs each calendar script
  against them, checks the real Outlook state after each step, and prints PASS/FAIL lines
  and counts only. Nothing is sent: meeting creation, forwarding and responding run in
  preview or on a hold (which must be refused). Everything it creates is deleted at the end.
  -Exe powershell runs every script under Windows PowerShell 5.1.
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
# Force the confirm gate on for the whole run: with it cleared, config.json could turn it off
# and a "preview" step would really send. Restored at the end.
$prevConfirm = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"

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
# Fails the run at once (after cleanup) if any item of this run reached Outbox or Sent.
function Assert-NothingSent([string]$prefix) {
    $n = Get-SentTestItems $prefix
    if ($n -gt 0) { $script:fails++; Write-Host "FAIL  SOMETHING WAS SENT ($n item(s) in Outbox/Sent). Stopping." -ForegroundColor Red; throw "sent" }
}
function J($r) { try { return ($r.Out | ConvertFrom-Json) } catch { return $null } }
function Find-Test([datetime]$from, [datetime]$to) {
    $items = $Namespace.GetDefaultFolder(9).Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $f = "[Start] >= '$($from.ToString('MM/dd/yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture))' AND [Start] < '$($to.ToString('MM/dd/yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture))'"
    return @(foreach ($e in $items.Restrict($f)) { if ($e.Subject -like "``[suite-test``]*") { $e } })
}

$tag = "[suite-test] cal $(Get-Date -Format 'MMddHHmmss')"
$day = (Get-Date).Date.AddDays(60)
while ($day.DayOfWeek -ne "Monday") { $day = $day.AddDays(1) }
$t6 = $day.AddHours(6)
$created = New-Object System.Collections.Generic.List[string]
Write-Host "Running calendar scripts under $Exe" -ForegroundColor Gray

try {
    # List (read only)
    $r = Run "outlook-calendar-list.ps1" @("-Days", "30", "-Limit", "500", "-Json"); $j = J $r
    Check "list 30d -Json parses, every event has an EntryID" ($r.Exit -eq 0 -and $j -and @($j.Events | Where-Object { -not $_.EntryID }).Count -eq 0) "exit=$($r.Exit)"
    Write-Host "      list 30d: $($j.Total) events, $($r.Ms) ms" -ForegroundColor Gray
    $r = Run "outlook-calendar-list.ps1" @("-Days", "7")
    Check "list 7d text" ($r.Exit -eq 0 -and $r.Out -match "--- Total: \d+ events")

    # Create a hold
    $r = Run "outlook-calendar-create.ps1" @("-Subject", "$tag hold", "-Start", $t6.ToString("yyyy-MM-dd HH:mm"), "-Duration", "30", "-BusyStatus", "Tentative", "-NoReminder", "-Json"); $j = J $r
    Check "create hold" ($r.Exit -eq 0 -and $j.EntryID -and $j.Kind -eq "hold" -and $j.Busy -eq "Tentative") "exit=$($r.Exit)"
    $hold = $j.EntryID; if ($hold) { $created.Add($hold) }
    $it = $Namespace.GetItemFromID($hold)
    Check "hold saved as tentative, no reminder, 30 min" ($it.BusyStatus -eq 1 -and -not $it.ReminderSet -and $it.Duration -eq 30)

    # Update it
    $r = Run "outlook-calendar-update.ps1" @("-EntryID", $hold, "-Start", $t6.AddMinutes(30).ToString("yyyy-MM-dd HH:mm"), "-Location", "suite-test room", "-BusyStatus", "Busy", "-Json"); $j = J $r
    $it = $Namespace.GetItemFromID($hold)
    Check "update moves start, keeps duration, sets location and show-as" ($r.Exit -eq 0 -and $it.Start -eq $t6.AddMinutes(30) -and $it.Duration -eq 30 -and $it.Location -eq "suite-test room" -and $it.BusyStatus -eq 2)
    $r = Run "outlook-calendar-update.ps1" @("-EntryID", $hold, "-Subject", "$tag hold2", "-SendUpdate")
    Check "update -SendUpdate (no -Confirm) on a hold is refused" ($r.Exit -eq 1); Assert-NothingSent $tag

    # It shows in the list as a hold
    $r = Run "outlook-calendar-list.ps1" @("-From", $day.ToString("yyyy-MM-dd"), "-To", $day.ToString("yyyy-MM-dd"), "-Json"); $j = J $r
    $mine = @($j.Events | Where-Object { $_.EntryID -eq $hold })
    Check "list -From/-To finds the hold, Kind=hold" ($mine.Count -eq 1 -and $mine[0].Kind -eq "hold")

    # Send-capable scripts run in PREVIEW mode only: never -Confirm, never a deliverable
    # address (".invalid" is a reserved TLD that can never be delivered). After every
    # preview, Assert-NothingSent stops the test at once if anything reached Outbox or Sent.
    $fake = "preview-only@suite-test.invalid"
    $r = Run "outlook-calendar-attendees.ps1" @("-EntryID", $hold, "-Json"); $j = J $r
    Check "attendees on a hold: 0 people" ($r.Exit -eq 0 -and @($j.People).Count -eq 0)
    $r = Run "outlook-calendar-respond.ps1" @("-EntryID", $hold, "-Response", "Accept")
    Check "respond to a hold is refused" ($r.Exit -eq 1); Assert-NothingSent $tag
    $r = Run "outlook-calendar-forward.ps1" @("-EntryID", $hold, "-To", $fake, "-Json"); $j = J $r
    Check "forward without -Send saves a draft" ($r.Exit -eq 0 -and $j.DraftEntryID); Assert-NothingSent $tag
    if ($j.DraftEntryID) { try { $Namespace.GetItemFromID($j.DraftEntryID).Delete() } catch { } }

    # Hold to invite: preview only, the hold must stay untouched
    $r = Run "outlook-calendar-hold-to-invite.ps1" @("-EntryID", $hold, "-Attendees", $fake, "-Subject", "$tag renamed")
    $it = $Namespace.GetItemFromID($hold)
    Check "hold-to-invite without -Confirm: preview, hold unchanged" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY" -and $it.MeetingStatus -eq 0 -and $it.Recipients.Count -eq 0 -and $it.Subject -notlike "*renamed"); Assert-NothingSent $tag
    $r = Run "outlook-calendar-hold-to-invite.ps1" @("-EntryID", $hold, "-Attendees", $fake, "-Json"); $j = J $r
    $it = $Namespace.GetItemFromID($hold)
    Check "hold-to-invite -Json preview is JSON, hold unchanged" ($r.Exit -eq 0 -and $j.Result -eq "preview" -and -not $j.Done -and $it.MeetingStatus -eq 0 -and $it.Recipients.Count -eq 0); Assert-NothingSent $tag
    $r = Run "outlook-calendar-hold-to-invite.ps1" @("-EntryID", $hold, "-Attendees", "zz-no-such-person-zz")
    $it = $Namespace.GetItemFromID($hold)
    Check "hold-to-invite with an unknown attendee fails, hold unchanged" ($r.Exit -eq 1 -and $it.MeetingStatus -eq 0); Assert-NothingSent $tag

    # The gate helper itself must return exactly one [bool] in every mode
    . (Join-Path $ScriptsDir "_calendar.ps1")
    foreach ($jm in $false, $true) {
        $res = @(Confirm-CalendarAction -Action "send" -Title "UNIT" -Item $it -Confirmed $false -What "unit test" -Json:$jm 6>$null)
        Check "Confirm-CalendarAction (-Json:$jm) returns exactly one [bool] `$false" ($res.Count -eq 1 -and $res[0] -is [bool] -and $res[0] -eq $false) "count=$($res.Count)"
    }

    # Week overview sees the hold
    $r = Run "outlook-calendar-week.ps1" @("-From", $day.ToString("yyyy-MM-dd"), "-Detail", "-Json"); $j = J $r
    $wd = @($j.Days | Where-Object { ([datetimeoffset]::Parse($_.Date)).LocalDateTime -eq $day })[0]
    Check "week -Detail lists the hold on its day" ($r.Exit -eq 0 -and @($wd.Events | Where-Object { $_.EntryID -eq $hold -and $_.Kind -eq "hold" }).Count -eq 1)

    # Meeting creation previews and saves nothing
    $r = Run "outlook-calendar-create.ps1" @("-Subject", "$tag meeting", "-Start", $t6.AddHours(2).ToString("yyyy-MM-dd HH:mm"), "-Attendees", $fake)
    Check "create with attendees, no -Confirm: preview" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY"); Assert-NothingSent $tag
    $r = Run "outlook-calendar-create.ps1" @("-Subject", "$tag meeting", "-Start", $t6.AddHours(2).ToString("yyyy-MM-dd HH:mm"), "-Attendees", $fake, "-Json"); $j = J $r
    Check "create with attendees -Json: JSON preview" ($r.Exit -eq 0 -and $j.Result -eq "preview"); Assert-NothingSent $tag
    Check "  and nothing was saved" (@(Find-Test $t6.AddHours(2) $t6.AddHours(3)).Count -eq 0)
    $r = Run "outlook-calendar-create.ps1" @("-Subject", "$tag meeting", "-Start", $t6.AddHours(2).ToString("yyyy-MM-dd HH:mm"), "-Attendees", "zz-no-such-person-zz")
    Check "create with an unknown attendee fails before saving" ($r.Exit -eq 1)

    # Recurring hold, delete one occurrence, then the series
    $r = Run "outlook-calendar-recurring.ps1" @("-Subject", "$tag series", "-Start", $t6.AddDays(1).ToString("yyyy-MM-dd HH:mm"), "-Pattern", "Daily", "-Occurrences", "3", "-Duration", "15", "-NoReminder", "-Json"); $j = J $r
    Check "recurring hold, 3 occurrences" ($r.Exit -eq 0 -and $j.EntryID -and $j.Recurring)
    $series = $j.EntryID; if ($series) { $created.Add($series) }
    $r = Run "outlook-calendar-list.ps1" @("-From", $day.AddDays(1).ToString("yyyy-MM-dd"), "-To", $day.AddDays(3).ToString("yyyy-MM-dd"), "-Json"); $j = J $r
    $occ = @($j.Events | Where-Object { $_.EntryID -eq $series })
    Check "list expands the series (3 occurrences, series EntryID)" ($occ.Count -eq 3)
    $r = Run "outlook-calendar-delete.ps1" @("-EntryID", $series, "-OccurrenceStart", $t6.AddDays(2).ToString("yyyy-MM-dd HH:mm"))
    Check "delete occurrence without -Confirm: preview" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY" -and $r.Out -match "occurrence only")
    $r = Run "outlook-calendar-delete.ps1" @("-EntryID", $series, "-OccurrenceStart", $t6.AddDays(2).ToString("yyyy-MM-dd HH:mm"), "-Confirm")
    Check "delete one occurrence" ($r.Exit -eq 0 -and @(Find-Test $day.AddDays(1) $day.AddDays(4)).Count -eq 2)
    $r = Run "outlook-calendar-delete.ps1" @("-EntryID", $series)
    Check "delete series preview warns ENTIRE SERIES" ($r.Out -match "ENTIRE SERIES")

    # Free/busy for self (read only)
    $me = $Namespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress   # read-only free/busy lookup
    $r = Run "outlook-freebusy.ps1" @("-Attendees", $me, "-Start", $day.AddHours(9).ToString("yyyy-MM-dd HH:mm"), "-End", $day.AddHours(17).ToString("yyyy-MM-dd HH:mm"), "-Json"); $j = J $r
    Check "freebusy self: 16 half-hour slots" ($r.Exit -eq 0 -and @($j.People)[0].Status.Length -eq 16)

    # Rooms (read only)
    $r = Run "outlook-calendar-rooms.ps1" @("-Start", $day.AddHours(10).ToString("yyyy-MM-dd HH:mm"), "-Limit", "20", "-Json"); $j = J $r
    Check "rooms -Json" ($r.Exit -eq 0 -or $r.Out -match "No room list") "exit=$($r.Exit)"
    if ($j) { Write-Host "      rooms: $(@($j.Rooms).Count) checked, $($j.Free) free, $($r.Ms) ms" -ForegroundColor Gray }

    # Delete the hold with -Confirm
    $r = Run "outlook-calendar-delete.ps1" @("-EntryID", $hold, "-Confirm")
    Check "delete hold" ($r.Exit -eq 0 -and @(Find-Test $t6 $t6.AddHours(1)).Count -eq 0)
    $r = Run "outlook-calendar-delete.ps1" @("-EntryID", $series, "-Confirm")
    Check "delete series" ($r.Exit -eq 0 -and @(Find-Test $day.AddDays(1) $day.AddDays(4)).Count -eq 0)
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
