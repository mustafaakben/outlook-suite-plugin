param(
    [datetime]$From,
    [switch]$Next,
    [int]$Weeks = 1,
    [int]$MaxPerDay = 2,
    [switch]$IncludeWeekends,
    [switch]$Detail,
    [string[]]$FirmHolds,
    [switch]$Json
)

# Outlook Calendar Week: a day-by-day overview of meetings vs holds vs all-day events,
# with each day's tier (open, holds-only, one-meeting, busy), away days, days over the
# daily meeting cap, and conflicting meetings.
# Usage: .\outlook-calendar-week.ps1                 # this week (Monday to Friday)
# Next week: .\outlook-calendar-week.ps1 -Next
# The week of a date, with every event listed: .\outlook-calendar-week.ps1 -From "2026-11-02" -Detail
# Two weeks, weekends too: .\outlook-calendar-week.ps1 -Weeks 2 -IncludeWeekends

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    $anchor = if ($PSBoundParameters.ContainsKey("From")) { $From.Date } else { (Get-Date).Date }
    $monday = $anchor.AddDays(-(([int]$anchor.DayOfWeek + 6) % 7))
    if ($Next) { $monday = $monday.AddDays(7) }
    $end = $monday.AddDays(7 * [Math]::Max(1, $Weeks))

    $events = Get-CalendarEvents -Namespace $Namespace -From $monday -To $end -FirmPatterns (Get-FirmHoldPatterns -Override $FirmHolds)
    Set-CalendarConflicts $events

    $dayList = New-Object System.Collections.Generic.List[object]
    for ($d = $monday; $d -lt $end; $d = $d.AddDays(1)) {
        $info = Get-CalendarDay -Records $events -Date $d -MaxPerDay $MaxPerDay
        $weekend = ($d.DayOfWeek -eq "Saturday" -or $d.DayOfWeek -eq "Sunday")
        if ($weekend -and -not $IncludeWeekends -and $info.Events.Count -eq 0) { continue }
        $timed = @($info.Events | Where-Object { $_.Kind -eq "meeting" -and $_.Busy -ne "Free" } | Sort-Object { $_.Start })
        $info.FirstMeeting = $null; $info.LastMeetingEnd = $null
        if ($timed.Count) { $info.FirstMeeting = $timed[0].Start; $info.LastMeetingEnd = ($timed | Sort-Object { $_.End })[-1].End }
        $mins = 0; foreach ($m in $timed) { $mins += ($m.End - $m.Start).TotalMinutes }
        $info.MeetingHours = [Math]::Round($mins / 60, 1)
        $dayList.Add($info)
    }
    $sumMeetings = 0; $sumHolds = 0
    foreach ($d in $dayList) { $sumMeetings += $d.Meetings; $sumHolds += $d.Holds }
    $totals = [ordered]@{
        Meetings  = $sumMeetings
        Holds     = $sumHolds
        OpenDays  = @($dayList | Where-Object { $_.Tier -eq 1 -and -not $_.Away }).Count
        AwayDays  = @($dayList | Where-Object { $_.Away }).Count
        OverCap   = @($dayList | Where-Object { $_.OverCap }).Count
        Conflicts = @($events | Where-Object { $_.ConflictsWith.Count -gt 0 }).Count
    }

    if ($Json) {
        if ($Detail) { Add-CalendarEntryIds -Namespace $Namespace -Records $events }
        $o = [ordered]@{ From = Format-MailDateIso $monday; To = Format-MailDateIso $end; MaxPerDay = $MaxPerDay }
        $o.Totals = $totals
        $o.Days = @($dayList | ForEach-Object {
            $j = ConvertTo-EventJson $_
            if ($Detail) { $j.Events = @($_.Events | ForEach-Object { ConvertTo-EventJson $_ }) }
            $j
        })
        Write-Json $o
        exit 0
    }

    Write-Host "`n=== WEEK: $($monday.ToString('ddd MMM d')) to $($end.AddDays(-1).ToString('ddd MMM d')) ===" -ForegroundColor Cyan
    Write-Host "$($totals.Meetings) meetings, $($totals.Holds) holds, $($totals.OpenDays) open days, $($totals.AwayDays) away, $($totals.OverCap) over the cap of $MaxPerDay, $($totals.Conflicts) meetings in conflict" -ForegroundColor Gray
    foreach ($d in $dayList) {
        $label = if ($d.Away) { "Away" } else { $d.TierName }
        $color = @{ 1 = "Green"; 2 = "Green"; 3 = "Yellow"; 4 = "Red" }[$d.Tier]
        if ($d.Away) { $color = "Magenta" }
        $span = ""
        if ($d.FirstMeeting) { $span = "meetings $($d.FirstMeeting.ToString('h:mm tt')) to $($d.LastMeetingEnd.ToString('h:mm tt')), $($d.MeetingHours) h" }
        if ($d.Firm) { $span = "$($d.Firm) firm  $span" }
        Write-Host ""
        Write-Host ("{0,-11} {1,-16} {2} meeting{3}, {4} hold{5}  {6}" -f $d.Date.ToString("ddd MMM d"), $label, $d.Meetings, $(if ($d.Meetings -ne 1) { "s" } else { "" }), $d.Holds, $(if ($d.Holds -ne 1) { "s" } else { "" }), $span) -ForegroundColor $color
        if ($d.AllDay.Count) { Write-Host "            all-day: $($d.AllDay -join '; ')" -ForegroundColor Gray }
        if ($d.OverCap) { Write-Host "            over the daily cap of $MaxPerDay meetings" -ForegroundColor Red }
        if ($d.Conflicts) { Write-Host "            $($d.Conflicts) meetings in conflict" -ForegroundColor Red }
        if ($Detail) {
            foreach ($e in @($d.Events | Where-Object { -not $_.AllDay } | Sort-Object { $_.Start })) {
                $flag = if ($e.ConflictsWith.Count) { "  CONFLICT" } else { "" }
                $c = if ($e.Kind -in @("meeting", "firm")) { "White" } else { "DarkGray" }
                Write-Host ("            {0,-8} {1,-8} {2,-8} {3}{4}" -f $e.Start.ToString("h:mm tt"), $e.End.ToString("h:mm tt"), $e.Kind, $e.Subject, $flag) -ForegroundColor $c
            }
        }
    }
    Write-Host "`nFind a slot: outlook-calendar-free.ps1 -From '$($monday.ToString('yyyy-MM-dd'))' -Days $((($end - $monday).TotalDays))" -ForegroundColor Gray
}
catch {
    Stop-WithError $_
}
