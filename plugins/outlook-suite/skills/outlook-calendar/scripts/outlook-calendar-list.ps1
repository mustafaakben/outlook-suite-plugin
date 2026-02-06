param(
    [int]$Days = 7,
    [switch]$Past,
    [int]$Limit = 20
)

# Outlook Calendar List Script
# Usage: .\outlook-calendar-list.ps1 -Days 7
# Include past events: .\outlook-calendar-list.ps1 -Days 7 -Past
# Limit results: .\outlook-calendar-list.ps1 -Days 14 -Limit 10

$Outlook = $null
$Namespace = $null
$Calendar = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")
    $Calendar = $Namespace.GetDefaultFolder(9)  # olFolderCalendar = 9

    $items = $Calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true

    # Set date range (fixed Outlook Restrict format)
    $now = Get-Date
    $restrictCulture = [System.Globalization.CultureInfo]::InvariantCulture
    if ($Past) {
        $startDate = $now.AddDays(-$Days).ToString("MM/dd/yyyy HH:mm", $restrictCulture)
        $endDate = $now.AddDays($Days).ToString("MM/dd/yyyy HH:mm", $restrictCulture)
        $rangeDesc = "Past $Days days and next $Days days"
    } else {
        $startDate = $now.ToString("MM/dd/yyyy HH:mm", $restrictCulture)
        $endDate = $now.AddDays($Days).ToString("MM/dd/yyyy HH:mm", $restrictCulture)
        $rangeDesc = "Next $Days days"
    }

    $filter = "[Start] >= '$startDate' AND [Start] <= '$endDate'"
    $events = $items.Restrict($filter)

    Write-Host "`n=== CALENDAR EVENTS ===" -ForegroundColor Cyan
    Write-Host "($rangeDesc)" -ForegroundColor Gray
    Write-Host ""

    $count = 0
    foreach ($event in $events) {
        if ($count -ge $Limit) { break }
        $count++

        # Format date/time
        $startTime = $event.Start
        $endTime = $event.End
        $duration = ($endTime - $startTime).TotalMinutes

        # Determine if it's today, tomorrow, or another day
        $dayLabel = ""
        if ($startTime.Date -eq $now.Date) {
            $dayLabel = "(Today)"
        } elseif ($startTime.Date -eq $now.Date.AddDays(1)) {
            $dayLabel = "(Tomorrow)"
        } elseif ($startTime.Date -eq $now.Date.AddDays(-1)) {
            $dayLabel = "(Yesterday)"
        }

        # All-day event formatting
        if ($event.AllDayEvent) {
            $timeStr = "All Day"
        } else {
            $timeStr = "$($startTime.ToString('h:mm tt')) - $($endTime.ToString('h:mm tt'))"
        }

        # Display event
        Write-Host "$count. $($event.Subject)" -ForegroundColor Yellow
        Write-Host "      EntryID: $($event.EntryID)" -ForegroundColor DarkGray
        Write-Host "      Date: $($startTime.ToString('ddd, MMM d, yyyy')) $dayLabel" -ForegroundColor White
        Write-Host "      Time: $timeStr" -ForegroundColor Gray

        if ($event.Location) {
            Write-Host "      Location: $($event.Location)" -ForegroundColor Gray
        }

        # Show organizer for meetings
        if ($event.MeetingStatus -ne 0) {
            Write-Host "      Organizer: $($event.Organizer)" -ForegroundColor Gray
        }

        # Show busy status
        $busyStatus = switch ($event.BusyStatus) {
            0 { "Free" }
            1 { "Tentative" }
            2 { "Busy" }
            3 { "Out of Office" }
            4 { "Working Elsewhere" }
            default { "Unknown" }
        }
        if ($busyStatus -ne "Busy") {
            Write-Host "      Status: $busyStatus" -ForegroundColor Gray
        }

        Write-Host ""
    }

    if ($count -eq 0) {
        Write-Host "No events found in this time range." -ForegroundColor Gray
    } else {
        Write-Host "--- Total: $count events ---" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Calendar) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Calendar) | Out-Null } catch {}
    }
    if ($Namespace) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
