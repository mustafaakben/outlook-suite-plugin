param(
    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [Parameter(Mandatory=$true)]
    [datetime]$Start,

    [datetime]$End,
    [string]$Location = "",
    [string]$Body = "",
    [switch]$AllDay,
    [int]$Reminder = 15,
    [string]$Attendees = ""
)

# Outlook Calendar Create Script
# Usage: .\outlook-calendar-create.ps1 -Subject "Meeting" -Start "2026-02-05 10:00"
# With end time: .\outlook-calendar-create.ps1 -Subject "Meeting" -Start "2026-02-05 10:00" -End "2026-02-05 11:00"
# All-day event: .\outlook-calendar-create.ps1 -Subject "Vacation" -Start "2026-02-10" -AllDay
# With attendees: .\outlook-calendar-create.ps1 -Subject "Team Meeting" -Start "2026-02-05 14:00" -Attendees "john@example.com; jane@example.com"

$Outlook = $null
$appointment = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    # Create appointment item (olAppointmentItem = 1)
    $appointment = $Outlook.CreateItem(1)

    $appointment.Subject = $Subject
    $appointment.Start = $Start

    # Set end time (default: 1 hour after start)
    if ($End) {
        $appointment.End = $End
    } else {
        if ($AllDay) {
            $appointment.End = $Start.AddDays(1)
        } else {
            $appointment.End = $Start.AddHours(1)
        }
    }

    # Set optional properties
    if ($Location) {
        $appointment.Location = $Location
    }

    if ($Body) {
        $appointment.Body = $Body
    }

    if ($AllDay) {
        $appointment.AllDayEvent = $true
    }

    # Set reminder
    $appointment.ReminderSet = $true
    $appointment.ReminderMinutesBeforeStart = $Reminder

    # If attendees specified, make it a meeting
    if ($Attendees) {
        $appointment.MeetingStatus = 1  # olMeeting
        $appointment.RequiredAttendees = $Attendees

        # Save first, then send invites
        $appointment.Save()
        $appointment.Send()

        Write-Host "`n=== MEETING CREATED & INVITES SENT ===" -ForegroundColor Green
        Write-Host "Attendees: $Attendees" -ForegroundColor Cyan
    } else {
        $appointment.Save()
        Write-Host "`n=== APPOINTMENT CREATED ===" -ForegroundColor Green
    }

    Write-Host "Subject: $Subject" -ForegroundColor Yellow
    Write-Host "Date: $($Start.ToString('ddd, MMM d, yyyy'))"

    if ($AllDay) {
        Write-Host "Time: All Day" -ForegroundColor Gray
    } else {
        Write-Host "Time: $($Start.ToString('h:mm tt')) - $($appointment.End.ToString('h:mm tt'))" -ForegroundColor Gray
    }

    if ($Location) {
        Write-Host "Location: $Location" -ForegroundColor Gray
    }

    Write-Host "Reminder: $Reminder minutes before" -ForegroundColor Gray
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($appointment) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($appointment) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
