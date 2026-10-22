param(
    [string]$Attendees = "",
    [int]$Duration = 30,
    [datetime]$From,
    [datetime]$To,
    [int]$Days = 7,
    [string]$WorkStart = "09:00",
    [string]$WorkEnd = "17:00",
    [ValidateSet(5, 10, 15, 30, 60)]
    [int]$Step = 30,
    [int]$Buffer = 0,
    [int]$MaxPerDay = 2,
    [int]$PerDay = 2,
    [int]$Top = 12,
    [switch]$IncludeWeekends,
    [string[]]$FirmHolds,
    [switch]$Propose,
    [int]$Count = 3,
    [switch]$Html,
    [string]$TimeZone = "",
    [switch]$Json
)

# Outlook Calendar Free: find meeting times for the user alone or with other people.
#
# It first summarizes the week (meetings, holds, firm events, all-day events, away days),
# then ranks each day and fills slots from the best tier down:
#   1. Open day         no meetings and no holds
#   2. Holds-only day   only holds (events with no one else invited); a slot may go over a hold.
#                       Firm events (classes, office hours, exams, travel, appointments, Out of
#                       Office, private; see -FirmHolds) are never booked over.
#   3. One-meeting day  one meeting booked, so the new one is the second (the daily maximum)
#   4. Busy day         the daily cap is already reached; offered only when tiers 1 to 3 have nothing
# Away days (all-day Out of Office or Busy) get no slots. Other people's free/busy comes
# from Exchange; tentative time is allowed but noted, and anyone whose free/busy could not
# be read is reported as NOT CHECKED. If nothing fits, the search widens by a week (twice
# at most) and then explains which days were skipped and why.
#
# Usage: .\outlook-calendar-free.ps1                                   # 30 min, next 7 days, just the user
# With people: .\outlook-calendar-free.ps1 -Attendees "Jane Doe; john@example.com" -Duration 60
# A given week: .\outlook-calendar-free.ps1 -From "2026-11-02" -To "2026-11-06" -Duration 45
# Hours and padding: .\outlook-calendar-free.ps1 -WorkStart 10 -WorkEnd "4:00 PM" -Buffer 15
# Email-ready options: .\outlook-calendar-free.ps1 -Attendees "Jane Doe" -Propose            # 3 lines
#                      .\outlook-calendar-free.ps1 -Propose -Count 4 -Html                   # <div><ul> for a draft
# Another time zone (display only): .\outlook-calendar-free.ps1 -Propose -TimeZone Pacific
# JSON: add -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    Set-DisplayTimeZone $TimeZone
    $now = Get-Date
    $ws = ConvertTo-TimeOfDay $WorkStart "-WorkStart"; $we = ConvertTo-TimeOfDay $WorkEnd "-WorkEnd"
    if ($we -le $ws) { throw "-WorkEnd must be after -WorkStart." }
    if ($Duration -lt 5) { throw "-Duration must be at least 5 minutes." }
    $first = if ($PSBoundParameters.ContainsKey("From")) { $From.Date } else { $now.Date }
    $last = if ($PSBoundParameters.ContainsKey("To")) { $To.Date } else { $first.AddDays($Days - 1) }
    if ($last -lt $first) { throw "-To is before -From." }
    $notes = New-Object System.Collections.Generic.List[string]
    if ($first -lt $now.Date) {
        if ($last -lt $now.Date) { throw "The whole range is in the past." }
        $notes.Add("Days before today were skipped.")
        $first = $now.Date
    }
    if (($last - $first).TotalDays -gt 27) { throw "Search at most 4 weeks at a time (Exchange free/busy covers about 29 days)." }
    $limit = $first.AddDays(27)

    # Other people: one free/busy string each, starting at midnight of the first day, one
    # character per $Step minutes of ELAPSED time (so it stays aligned across a DST change;
    # index with UTC differences, never local wall-clock differences).
    $baseUtc = $first.ToUniversalTime()
    $people = New-Object System.Collections.Generic.List[object]
    foreach ($text in (Split-AttendeeList $Attendees)) {
        $a = Resolve-CalendarAttendee -Namespace $Namespace -Text $text
        $label = if ($a.Name) { $a.Name } else { $text }
        $p = [ordered]@{ Input = $text; Name = $a.Name; Label = $label; Email = $a.Email; Resolved = $a.Resolved; Note = $a.Note; FreeBusy = ""; DataUntil = $null }
        if ($a.Resolved) {
            try { $p.FreeBusy = [string]$a.Recipient.FreeBusy($first, $Step, $true) }
            catch { $p.Note = "No free/busy data (calendar not shared?): $($_.Exception.Message)" }
            if (-not $p.FreeBusy -and -not $p.Note) { $p.Note = "No free/busy data." }
            if ($p.FreeBusy) { $p.DataUntil = $baseUtc.AddMinutes($p.FreeBusy.Length * $Step).ToLocalTime() }
        }
        $people.Add($p)
    }
    $notChecked = @($people | Where-Object { -not $_.FreeBusy } | ForEach-Object { $_.Label })

    # Slots for days [$rangeFrom, $rangeTo]; also fills $dayList and $skipped for the overview.
    $search = {
        param([datetime]$rangeFrom, [datetime]$rangeTo)
        $weekFrom = $rangeFrom.AddDays(-(([int]$rangeFrom.DayOfWeek + 6) % 7))
        $weekTo = $rangeTo.AddDays(7 - (([int]$rangeTo.DayOfWeek + 6) % 7))
        $events = Get-CalendarEvents -Namespace $Namespace -From $weekFrom -To $weekTo -FirmPatterns (Get-FirmHoldPatterns -Override $FirmHolds)
        Set-CalendarConflicts $events
        $script:dayList = New-Object System.Collections.Generic.List[object]
        $script:skipped = New-Object System.Collections.Generic.List[object]
        $found = New-Object System.Collections.Generic.List[object]
        for ($d = $weekFrom; $d -lt $weekTo; $d = $d.AddDays(1)) {
            $info = Get-CalendarDay -Records $events -Date $d -MaxPerDay $MaxPerDay
            $weekend = ($d.DayOfWeek -eq "Saturday" -or $d.DayOfWeek -eq "Sunday")
            $inRange = ($d -ge $rangeFrom -and $d -le $rangeTo)
            $info.InRange = $inRange
            $info.Weekend = $weekend
            if (-not ($weekend -and -not $IncludeWeekends -and ($info.Meetings + $info.Holds + $info.Firm + $info.AllDay.Count) -eq 0)) { $script:dayList.Add($info) }
            if (-not $inRange) { continue }
            if ($weekend -and -not $IncludeWeekends) { continue }
            if ($info.Away) { $script:skipped.Add([ordered]@{ Date = $d; Reason = "away (all-day: $($info.AllDay -join '; '))" }); continue }

            $winStart = $d.Add($ws); $winEnd = $d.Add($we)
            if ($d -eq $now.Date) {
                # Today: start at now rounded UP to the next step, never in the past.
                $mins = [Math]::Ceiling(($now - $d).TotalMinutes / $Step) * $Step
                $winStart = [datetime]::new([Math]::Max($winStart.Ticks, $d.AddMinutes($mins).Ticks))
            }
            $blocks = @($info.Events | Where-Object { $_.Blocking -and $_.Kind -in @("meeting", "firm") })
            $holds = @($info.Events | Where-Object { $_.Kind -eq "hold" -and $_.Busy -ne "Free" })

            $cands = New-Object System.Collections.Generic.List[object]
            for ($t = $winStart; $t.AddMinutes($Duration) -le $winEnd; $t = $t.AddMinutes($Step)) {
                $e = $t.AddMinutes($Duration)
                $busy = $false
                foreach ($b in $blocks) { if ($b.Start.AddMinutes(-$Buffer) -lt $e -and $t -lt $b.End.AddMinutes($Buffer)) { $busy = $true; break } }
                if ($busy) { continue }
                $tentative = @(); $unchecked = @($notChecked)
                $i0 = [int][Math]::Floor(($t.ToUniversalTime() - $baseUtc).TotalMinutes / $Step)
                $i1 = [int][Math]::Ceiling(($e.ToUniversalTime() - $baseUtc).TotalMinutes / $Step)
                foreach ($p in $people) {
                    if (-not $p.FreeBusy) { continue }
                    if ($i0 -lt 0 -or $p.FreeBusy.Length -lt $i1) { $unchecked += "$($p.Label) (no data beyond $($p.DataUntil.ToString('MMM d')))"; continue }
                    $part = $p.FreeBusy.Substring($i0, $i1 - $i0)
                    if ($part -match "[23]") { $busy = $true; break }
                    if ($part -match "1") { $tentative += $p.Label }
                }
                if ($busy) { continue }
                $over = @($holds | Where-Object { $_.Start -lt $e -and $t -lt $_.End } | ForEach-Object { $_.Subject })
                $cands.Add([ordered]@{ Start = $t; End = $e; Tier = $info.Tier; TierName = $info.TierName; OverHold = $over; Tentative = $tentative; NotChecked = $unchecked })
            }
            if ($cands.Count -eq 0) { $script:skipped.Add([ordered]@{ Date = $d; Reason = "no room: meetings and firm events fill $($ws.ToString('hh\:mm')) to $($we.ToString('hh\:mm'))$(if ($people.Count) { ' (or attendees are busy)' })" }); continue }
            # Best few per day: fewest holds covered and tentative people first, then earliest.
            # First pass keeps picks at least 2 hours apart (real choices, not back-to-back slots).
            $ranked = @($cands | Sort-Object { $_.OverHold.Count }, { $_.Tentative.Count }, { $_.Start })
            $picked = @()
            foreach ($gap in 120, 0) {
                foreach ($c in $ranked) {
                    if ($picked.Count -ge $PerDay) { break }
                    if (@($picked | Where-Object { $c.Start -lt $_.End.AddMinutes($gap) -and $_.Start -lt $c.End.AddMinutes($gap) }).Count -gt 0) { continue }
                    $picked += $c
                }
            }
            foreach ($c in @($picked | Sort-Object { $_.Start })) { $found.Add($c) }
        }
        return , $found
    }

    $slots = & $search $first $last
    # Never "no time": widen by a week (twice at most) when nothing fits.
    $widened = 0
    while ($slots.Count -eq 0 -and $widened -lt 2 -and $last -lt $limit) {
        $newLast = $last.AddDays(7); if ($newLast -gt $limit) { $newLast = $limit }
        $notes.Add("Nothing fit by $($last.ToString('ddd MMM d')), so the search was widened to $($newLast.ToString('ddd MMM d')).")
        $last = $newLast; $widened++
        $slots = & $search $first $last
    }

    # Fill from the best tier down; busy days only when nothing better exists.
    $offer = @($slots | Where-Object { $_.Tier -le 3 } | Sort-Object { $_.Tier }, { $_.Start })
    $busyOnly = $false
    if ($offer.Count -eq 0) { $offer = @($slots | Sort-Object { $_.Start }); $busyOnly = ($offer.Count -gt 0) }
    $offer = @($offer | Select-Object -First $Top)

    # -Propose: the top -Count slots, best tier first, one per day before doubling up.
    $proposals = @()
    if ($Propose) {
        $byDay = @{}
        foreach ($pass in 1, 2) {
            foreach ($s in $offer) {
                if ($proposals.Count -ge $Count) { break }
                $k = $s.Start.Date
                if ($proposals -contains $s) { continue }
                if ($pass -eq 1 -and $byDay.ContainsKey($k)) { continue }
                $proposals += $s; $byDay[$k] = 1
            }
        }
        $proposals = @($proposals | Sort-Object { $_.Tier }, { $_.Start })
    }
    $style = "font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);"
    $htmlBlock = ""
    if ($Propose) {
        $items = ($proposals | ForEach-Object { "<li style=""$style"" class=""elementToProof"">$([System.Net.WebUtility]::HtmlEncode((Format-SlotLine $_.Start $_.End)))</li>" }) -join ""
        $htmlBlock = "<div style=""$style"" class=""elementToProof""><ul>$items</ul></div>"
    }

    $rangeText = "$($first.ToString('ddd MMM d')) to $($last.ToString('ddd MMM d'))"
    $hoursText = "$([datetime]::Today.Add($ws).ToString('h:mm tt')) to $([datetime]::Today.Add($we).ToString('h:mm tt'))"

    if ($Json) {
        $o = [ordered]@{ From = Format-MailDateIso $first; To = Format-MailDateIso $last.AddDays(1); Duration = $Duration }
        $o.WorkHours = "$($ws.ToString('hh\:mm'))-$($we.ToString('hh\:mm'))"; $o.MaxPerDay = $MaxPerDay; $o.BusyDaysOnly = $busyOnly
        if ($script:DisplayTz) { $o.TimeZoneShown = $script:DisplayTz.Id }
        $o.NotChecked = $notChecked
        $o.Notes = $notes.ToArray()
        $o.People = @($people | ForEach-Object { [ordered]@{ Input = $_.Input; Name = $_.Name; Email = $_.Email; Resolved = $_.Resolved; HasFreeBusy = [bool]$_.FreeBusy; DataUntil = Format-MailDateIso $_.DataUntil; Note = $_.Note } })
        $o.Week = @($script:dayList | ForEach-Object { ConvertTo-EventJson $_ })
        $o.Skipped = @($script:skipped | ForEach-Object { ConvertTo-EventJson $_ })
        $o.Slots = @($offer | ForEach-Object { $j = ConvertTo-EventJson $_; $j.Text = Format-SlotLine $_.Start $_.End; $j })
        if ($Propose) {
            $o.Proposals = @($proposals | ForEach-Object { $j = ConvertTo-EventJson $_; $j.Text = Format-SlotLine $_.Start $_.End; $j })
            $o.Html = $htmlBlock
        }
        Write-Json $o
        exit 0
    }

    if ($Propose) {
        if ($proposals.Count -eq 0) { Stop-WithError "No slot to propose in $rangeText." "Try a later -From, wider -WorkStart/-WorkEnd, a shorter -Duration, or -IncludeWeekends." }
        if ($Html) { Write-Output $htmlBlock; exit 0 }
        foreach ($s in $proposals) { Write-Output (Format-SlotLine $s.Start $s.End) }
        if ($notChecked.Count) { Write-Host "NOT CHECKED: $($notChecked -join ', ')" -ForegroundColor Yellow }
        exit 0
    }

    $who = "just you"
    if ($people.Count) { $who = "you + " + (($people | ForEach-Object { $_.Label }) -join ", ") }
    Write-Host "`n=== FREE TIME: $Duration min, $who ===" -ForegroundColor Cyan
    Write-Host "$rangeText, $hoursText$(if ($Buffer) { ", $Buffer min buffer" }), max $MaxPerDay meetings a day$(if ($script:DisplayTzLabel) { ", times shown in $script:DisplayTzLabel" })" -ForegroundColor Gray
    if ($notChecked.Count) { Write-Host "NOT CHECKED: $($notChecked -join ', ') (their calendars were not read; slots may clash for them)" -ForegroundColor Yellow }
    foreach ($p in $people) { if ($p.Note) { Write-Host "  ! $($p.Input): $($p.Note)" -ForegroundColor Yellow } }
    foreach ($n in $notes) { Write-Host "  $n" -ForegroundColor Yellow }

    Write-Host "`nWeek at a glance:" -ForegroundColor Cyan
    foreach ($d in $script:dayList) {
        $line = "  {0,-11} {1} meeting{2}, {3} hold{4}" -f $d.Date.ToString("ddd MMM d"), $d.Meetings, $(if ($d.Meetings -ne 1) { "s" } else { "" }), $d.Holds, $(if ($d.Holds -ne 1) { "s" } else { "" })
        $label = $d.TierName
        if ($d.Away) { $label = "Away" }
        $color = @{ 1 = "Green"; 2 = "Green"; 3 = "Yellow"; 4 = "Red" }[$d.Tier]
        if ($d.Away) { $color = "Magenta" }
        if (-not $d.InRange) { $color = "DarkGray" }
        $extra = @()
        if ($d.Firm) { $extra += "$($d.Firm) firm" }
        if ($d.AllDay.Count) { $extra += "all-day: " + ($d.AllDay -join "; ") }
        if ($d.OverCap) { $extra += "over the daily cap" }
        if ($d.Conflicts) { $extra += "$($d.Conflicts) in conflict" }
        Write-Host ("{0,-34} {1,-16} {2}" -f $line, $label, ($extra -join ", ")) -ForegroundColor $color
    }

    if ($offer.Count -eq 0) {
        Write-Host "`nEvery working day from $($first.ToString('MMM d')) to $($last.ToString('MMM d')) is taken:" -ForegroundColor Yellow
        foreach ($s in $script:skipped) { Write-Host "  $($s.Date.ToString('ddd MMM d')): $($s.Reason)" -ForegroundColor Gray }
        Write-Host "Options: a later -From, wider -WorkStart/-WorkEnd, a shorter -Duration, -IncludeWeekends, or -MaxPerDay 3." -ForegroundColor Gray
        exit 0
    }
    if ($busyOnly) { Write-Host "`nNo open, holds-only or one-meeting day has room, so these are on busy days (over the daily cap)." -ForegroundColor Yellow }
    foreach ($tier in 1..4) {
        $group = @($offer | Where-Object { $_.Tier -eq $tier })
        if ($group.Count -eq 0) { continue }
        Write-Host "`n$($script:TierNames[$tier])s:" -ForegroundColor Cyan
        foreach ($s in $group) {
            $n = @()
            if ($s.OverHold.Count) { $n += "over hold: " + ($s.OverHold -join "; ") }
            if ($s.Tentative.Count) { $n += "tentative: " + ($s.Tentative -join ", ") }
            if ($s.NotChecked.Count) { $n += "NOT CHECKED: " + ($s.NotChecked -join ", ") }
            $tail = if ($n.Count) { "  ($($n -join '; '))" } else { "" }
            Write-Host "  $(Format-SlotLine $s.Start $s.End)$tail"
        }
    }
    Write-Host "`nBook one: outlook-calendar-create.ps1 -Subject ... -Start '<start>' -Duration $Duration$(if ($people.Count) { " -Attendees '...' -Confirm" }). Email-ready list: add -Propose (or -Propose -Html)." -ForegroundColor Gray
}
catch {
    Stop-WithError $_
}
