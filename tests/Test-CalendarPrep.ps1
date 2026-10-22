<#
.SYNOPSIS
  Live test for outlook-calendar-prep.ps1 (read-only) and outlook-calendar-cancel.ps1
  (PREVIEW ONLY: never -Confirm).

.DESCRIPTION
  Prep: -Next and -EntryID return the invite, people with directory details and mail counts,
  and related mail as subjects and dates only (no bodies), within about 3 s.
  Cancel: refused on a "[suite-test]" hold; on one real meeting the user organizes it only
  previews, and the test proves nothing changed: the item's MeetingStatus, recipients,
  subject and body are unchanged and no new item with that subject reached Outbox or Sent.
  Prints PASS/FAIL, counts and timings only.
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
function J($r) { try { return ($r.Out | ConvertFrom-Json) } catch { return $null } }
# Items in Outbox + Sent whose subject contains $text, counted by a separate process.
function Get-OutgoingCount([string]$text) {
    $env:SUITE_TEST_QUERY = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%$($text.Replace("'", "''"))%'"
    return [int](& pwsh -NoProfile -Command '$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI"); $ns.GetDefaultFolder(4).Items.Restrict($env:SUITE_TEST_QUERY).Count + $ns.GetDefaultFolder(5).Items.Restrict($env:SUITE_TEST_QUERY).Count')
}

$tag = "[suite-test] prep $(Get-Date -Format 'MMddHHmmss')"
Write-Host "Running under $Exe" -ForegroundColor Gray
try {
    # Prep, read only
    $r = Run "outlook-calendar-prep.ps1" @("-Next", "-Json"); $j = J $r
    Check "prep -Next -Json" ($r.Exit -eq 0 -and $j -and $j.Kind -eq "meeting") "exit=$($r.Exit)"
    Write-Host "      prep -Next: $(@($j.People).Count) people, $(@($j.People | Where-Object { $_.Title }).Count) with title, $(@($j.People | Where-Object { $_.Received + $_.Sent }).Count) with mail, $(@($j.RelatedMail).Count) related, inner $($j.ElapsedMs) ms, total $($r.Ms) ms" -ForegroundColor Gray
    Check "prep stays near the 3 s target (inner work under 4 s)" ($j.ElapsedMs -lt 4000) "ms=$($j.ElapsedMs)"
    $keys = @($j.RelatedMail | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
    Check "related mail carries subjects and dates only, never bodies" (@($keys | Where-Object { $_ -notin @("Subject", "Date", "Who", "Folder", "Matches") }).Count -eq 0) "keys=$($keys -join ',')"
    Check "prep never lists the user among the people" (@($j.People | Where-Object { $_.Name -eq [string]$Namespace.CurrentUser.Name }).Count -eq 0)
    $r = Run "outlook-calendar-prep.ps1" @("-Next")
    Check "prep text output" ($r.Exit -eq 0 -and $r.Out -match "MEETING PREP" -and $r.Out -match "nothing was changed or sent")

    # Cancel: refused on a hold
    $day = (Get-Date).Date.AddDays(60)
    $c = J (Run "outlook-calendar-create.ps1" @("-Subject", "$tag hold", "-Start", $day.AddHours(6).ToString("yyyy-MM-dd HH:mm"), "-Duration", "30", "-NoReminder", "-Json"))
    $r = Run "outlook-calendar-cancel.ps1" @("-EntryID", $c.EntryID)
    Check "cancel on a hold is refused (nothing to notify)" ($r.Exit -eq 1 -and $r.Out -match "delete")
    Check "  and nothing reached Outbox or Sent" ((Get-OutgoingCount $tag) -eq 0)

    # Cancel: preview only on one real meeting the user organizes
    $l = J (Run "outlook-calendar-list.ps1" @("-Days", "30", "-Kind", "meeting", "-Limit", "200", "-Json"))
    $mine = @($l.Events | Where-Object { $_.Response -eq "Organizer" -and -not $_.Recurring })
    if ($mine.Count -eq 0) { $mine = @($l.Events | Where-Object { $_.Response -eq "Organizer" }) }
    if ($mine.Count -eq 0) { Write-Host "SKIP  cancel preview: no meeting organized by the user in the next 30 days" -ForegroundColor Yellow }
    else {
        $m = $mine[0]
        $before = $Namespace.GetItemFromID($m.EntryID)
        $snap = @([int]$before.MeetingStatus, [int]$before.Recipients.Count, [string]$before.Subject, ([string]$before.Body).Length)
        $outBefore = Get-OutgoingCount $snap[2]
        $r = Run "outlook-calendar-cancel.ps1" @("-EntryID", $m.EntryID, "-Message", "suite-test preview, never sent")
        Check "cancel without -Confirm prints a preview" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY" -and $r.Out -match "Notifies")
        $r = Run "outlook-calendar-cancel.ps1" @("-EntryID", $m.EntryID, "-Json"); $p = J $r
        Check "cancel -Json without -Confirm: one JSON preview object" ($r.Exit -eq 0 -and $p.Result -eq "preview" -and -not $p.Done)
        $after = $Namespace.GetItemFromID($m.EntryID)
        $now2 = @([int]$after.MeetingStatus, [int]$after.Recipients.Count, [string]$after.Subject, ([string]$after.Body).Length)
        Check "the meeting is unchanged (status, recipients, subject, body)" ((($snap -join "|") -eq ($now2 -join "|")) -and $now2[0] -eq 1)
        $outAfter = Get-OutgoingCount $snap[2]
        Check "no cancellation reached Outbox or Sent" ($outAfter -eq $outBefore) "before=$outBefore after=$outAfter"
        if ($outAfter -gt $outBefore) { throw "SOMETHING WAS SENT. Stopping." }
    }

    # Respond: preview only on one real invitation (a series if there is one)
    $inv = @($l.Events | Where-Object { $_.Response -ne "Organizer" -and $_.Recurring })
    if ($inv.Count -eq 0) { $inv = @($l.Events | Where-Object { $_.Response -ne "Organizer" }) }
    if ($inv.Count -eq 0) { Write-Host "SKIP  respond preview: no invitation in the next 30 days" -ForegroundColor Yellow }
    else {
        $m = $inv[0]
        $before = $Namespace.GetItemFromID($m.EntryID)
        $snap = @([int]$before.MeetingStatus, [int]$before.BusyStatus, [int]$before.ResponseStatus, [string]$before.Subject)
        $outBefore = Get-OutgoingCount $snap[3]
        $r = Run "outlook-calendar-respond.ps1" @("-EntryID", $m.EntryID, "-Response", "Accept", "-NoResponse", "-Json"); $p = J $r
        Check "respond -NoResponse preview shows old and new show-as" ($r.Exit -eq 0 -and $p.Result -eq "preview" -and $p.ShowAs -match " -> ")
        if ($m.Recurring) { Check "respond on a series EntryID: preview says ENTIRE SERIES" ($p.Scope -match "ENTIRE SERIES") }
        $r = Run "outlook-calendar-respond.ps1" @("-EntryID", $m.EntryID, "-Response", "Tentative")
        Check "respond without -Confirm prints a preview" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY")
        $after = $Namespace.GetItemFromID($m.EntryID)
        Check "the invitation is unchanged (status, show-as, response, subject)" ((($snap -join "|") -eq (@([int]$after.MeetingStatus, [int]$after.BusyStatus, [int]$after.ResponseStatus, [string]$after.Subject) -join "|")))
        $outAfter = Get-OutgoingCount $snap[3]
        Check "no response reached Outbox or Sent" ($outAfter -eq $outBefore) "before=$outBefore after=$outAfter"
        if ($outAfter -gt $outBefore) { throw "SOMETHING WAS SENT. Stopping." }
    }

    # The gate helper returns exactly one [bool] in both modes
    . (Join-Path $ScriptsDir "_calendar.ps1")
    $it = $Namespace.GetItemFromID($c.EntryID)
    foreach ($jm in $false, $true) {
        $res = @(Confirm-CalendarAction -Action "send" -Title "UNIT" -Item $it -Confirmed $false -What "unit test" -Json:$jm 6>$null)
        Check "Confirm-CalendarAction (-Json:$jm) returns exactly one [bool] `$false" ($res.Count -eq 1 -and $res[0] -is [bool] -and $res[0] -eq $false) "count=$($res.Count)"
    }
}
finally {
    Check "nothing was sent: 0 test items in Outbox or Sent (checked from a new process)" ((Get-OutgoingCount $tag) -eq 0)
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $prevConfirm
    $q = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%$($tag.Replace("'", "''"))%'"
    foreach ($fid in 9, 16, 3) { foreach ($i in @($Namespace.GetDefaultFolder($fid).Items.Restrict($q))) { try { $i.Delete() } catch { } } }
    $env:SUITE_TEST_QUERY = $q
    $left = [int](& pwsh -NoProfile -Command '$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI"); $n = 0; foreach ($fid in 9, 16, 3, 4, 5) { $n += $ns.GetDefaultFolder($fid).Items.Restrict($env:SUITE_TEST_QUERY).Count }; $n')
    Check "cleanup: no test items left (checked from a new process)" ($left -eq 0) "left=$left"
}
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green } else { Write-Host "$($script:fails) FAILED" -ForegroundColor Red; exit 1 }
