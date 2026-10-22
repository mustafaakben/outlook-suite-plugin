param(
    [Parameter(Mandatory=$true)]
    [string]$Attendees,

    [Parameter(Mandatory=$true)]
    [datetime]$Start,

    [datetime]$End,
    [ValidateRange(5, 1440)]
    [int]$IntervalMinutes = 30,
    [switch]$Json
)

# Outlook Free/Busy: a grid of when each person is free, tentative, busy or away.
# Names or emails; names are resolved through the directory. To find meeting slots,
# use outlook-calendar-free.ps1 instead.
# Usage: .\outlook-freebusy.ps1 -Attendees "john@example.com" -Start "2026-02-10 09:00"
# Multiple: .\outlook-freebusy.ps1 -Attendees "john@example.com; Jane Doe" -Start "2026-02-10 09:00" -End "2026-02-10 17:00"
# JSON (one status string per person, one character per interval): add -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

$symbols = @{ "0" = "  .   "; "1" = "  ?   "; "2" = " BUSY "; "3" = " OOF  "; "4" = " ELSE " }
$colors = @{ "0" = "Green"; "1" = "Yellow"; "2" = "Red"; "3" = "Magenta"; "4" = "Cyan" }

try {
    if (-not $End) { $End = $Start.AddHours(8) }
    if ($End -le $Start) { throw "End time must be later than start time." }
    $null = Connect-Outlook

    # FreeBusy strings start at midnight of the start date, one character per interval.
    $first = [int][Math]::Floor(($Start - $Start.Date).TotalMinutes / $IntervalMinutes)
    $slots = [int][Math]::Ceiling(($End - $Start).TotalMinutes / $IntervalMinutes)

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($text in (Split-AttendeeList $Attendees)) {
        $a = Resolve-CalendarAttendee -Namespace $Namespace -Text $text
        $row = [ordered]@{ Input = $text; Name = $a.Name; Email = $a.Email; Resolved = $a.Resolved; Status = ""; Note = $a.Note }
        if ($a.Resolved) {
            try {
                $fb = [string]$a.Recipient.FreeBusy($Start.Date, $IntervalMinutes, $true)
                $row.Status = $fb.Substring($first, [Math]::Min($slots, $fb.Length - $first))
            } catch { $row.Note = "No free/busy data: $($_.Exception.Message)" }
        }
        $rows.Add($row)
    }

    if ($Json) {
        # .ToArray(), not @($rows): @() over a List of ordered dictionaries fails in PowerShell 7
        # with "Argument types do not match".
        $isoStart = Format-MailDateIso $Start; $isoEnd = Format-MailDateIso $End
        $o = [ordered]@{ Start = $isoStart; End = $isoEnd; IntervalMinutes = $IntervalMinutes }
        $o.Legend = "0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere"
        $o.People = $rows.ToArray()
        Write-Json $o
        exit 0
    }

    Write-Host "`n=== FREE/BUSY: $($Start.ToString('ddd, MMM d h:mm tt')) - $($End.ToString('h:mm tt')) ($IntervalMinutes min) ===" -ForegroundColor Cyan
    $w = 24
    Write-Host ("{0,-$w}" -f "Person") -NoNewline -ForegroundColor Cyan
    Write-Host " | " -NoNewline
    for ($i = 0; $i -lt $slots; $i++) { Write-Host ("{0,-6}" -f $Start.AddMinutes($i * $IntervalMinutes).ToString("h:mm")) -NoNewline -ForegroundColor Gray }
    Write-Host ""
    Write-Host ("-" * ($w + 3 + $slots * 6)) -ForegroundColor Gray
    foreach ($r in $rows) {
        $label = if ($r.Name) { $r.Name } else { $r.Input }
        if ($label.Length -gt $w) { $label = $label.Substring(0, $w - 3) + "..." }
        Write-Host ("{0,-$w}" -f $label) -NoNewline -ForegroundColor White
        Write-Host " | " -NoNewline
        if (-not $r.Status) { Write-Host $r.Note -ForegroundColor Red; continue }
        foreach ($ch in $r.Status.ToCharArray()) {
            $k = [string]$ch; $c = $colors[$k]; if (-not $c) { $c = "Gray" }
            $s = $symbols[$k]; if (-not $s) { $s = "  ?   " }
            Write-Host $s -NoNewline -ForegroundColor $c
        }
        Write-Host ""
    }
    Write-Host "`nLegend:  .  free   ?  tentative   BUSY busy   OOF out of office   ELSE working elsewhere" -ForegroundColor Gray
}
catch {
    Stop-WithError $_
}
