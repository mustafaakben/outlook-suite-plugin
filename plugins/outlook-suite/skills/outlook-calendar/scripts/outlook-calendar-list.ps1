param(
    [int]$Days = 7,
    [switch]$Past,
    [int]$Limit = 20,
    [datetime]$From,
    [datetime]$To,
    [ValidateSet("", "meeting", "hold", "firm", "allday")]
    [string]$Kind = "",
    [switch]$ConflictsOnly,
    [string[]]$FirmHolds,
    [string]$TimeZone = "",
    [switch]$Json
)

# Outlook Calendar List: events in a range, recurrences expanded, each tagged
# meeting / hold / firm / allday, with overlapping meetings flagged as conflicts.
# Usage: .\outlook-calendar-list.ps1 -Days 7
# Include past events: .\outlook-calendar-list.ps1 -Days 7 -Past
# Exact range: .\outlook-calendar-list.ps1 -From "2026-10-26" -To "2026-10-30"
# Only meetings: .\outlook-calendar-list.ps1 -Days 14 -Kind meeting
# Only conflicts: .\outlook-calendar-list.ps1 -Days 14 -ConflictsOnly
# JSON: .\outlook-calendar-list.ps1 -Days 7 -Json
# Times shown in another zone (display only): .\outlook-calendar-list.ps1 -Days 7 -TimeZone Pacific

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    Set-DisplayTimeZone $TimeZone
    $now = Get-Date
    if ($From) {
        $rangeFrom = $From
        $rangeTo = if ($To) { $To.Date.AddDays(1) } else { $From.Date.AddDays($Days) }
        if ($To -and $To.TimeOfDay.TotalMinutes -gt 0) { $rangeTo = $To }
        $rangeDesc = "$($rangeFrom.ToString('ddd, MMM d')) to $($rangeTo.AddMinutes(-1).ToString('ddd, MMM d'))"
    } elseif ($Past) {
        $rangeFrom = $now.AddDays(-$Days); $rangeTo = $now.AddDays($Days)
        $rangeDesc = "Past $Days days and next $Days days"
    } else {
        $rangeFrom = $now; $rangeTo = $now.AddDays($Days)
        $rangeDesc = "Next $Days days"
    }

    $all = Get-CalendarEvents -Namespace $Namespace -From $rangeFrom -To $rangeTo -FirmPatterns (Get-FirmHoldPatterns -Override $FirmHolds)
    Set-CalendarConflicts $all
    $events = @($all | Where-Object { (-not $Kind -or $_.Kind -eq $Kind) -and (-not $ConflictsOnly -or $_.ConflictsWith.Count -gt 0) })
    $total = $events.Count
    $shown = @($events | Select-Object -First $Limit)
    Add-CalendarEntryIds -Namespace $Namespace -Records $shown

    if ($Json) {
        Write-Json ([ordered]@{
            Range  = [ordered]@{ From = Format-MailDateIso $rangeFrom; To = Format-MailDateIso $rangeTo }
            Total  = $total
            Shown  = $shown.Count
            Events = @($shown | ForEach-Object { ConvertTo-EventJson $_ })
        })
        exit 0
    }

    Write-Host "`n=== CALENDAR EVENTS ===" -ForegroundColor Cyan
    Write-Host "($rangeDesc)" -ForegroundColor Gray
    Write-Host ""
    $i = 0
    foreach ($e in $shown) {
        $i++
        Write-EventRecord $e -Prefix "$i. "
        Write-Host ""
    }
    if ($total -eq 0) {
        Write-Host "No events found in this time range." -ForegroundColor Gray
    } else {
        $m = @($events | Where-Object { $_.Kind -eq "meeting" }).Count
        $h = @($events | Where-Object { $_.Kind -eq "hold" }).Count
        $fm = @($events | Where-Object { $_.Kind -eq "firm" }).Count
        $a = @($events | Where-Object { $_.Kind -eq "allday" }).Count
        $c = @($events | Where-Object { $_.ConflictsWith.Count -gt 0 }).Count
        $more = if ($total -gt $shown.Count) { " (showing $($shown.Count); raise -Limit for more)" } else { "" }
        Write-Host "--- Total: $total events: $m meetings, $h holds, $fm firm, $a all-day, $c in conflict$more ---" -ForegroundColor Cyan
    }
}
catch {
    Stop-WithError $_
}
