param(
    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [Parameter(Mandatory=$true)]
    [datetime]$Start,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Daily", "Weekly", "Monthly", "Yearly")]
    [string]$Pattern,

    [int]$Interval = 1,
    [int]$Duration = 60,
    [string]$Location = "",
    [string]$Body = "",
    [string]$Attendees = "",
    [int]$Occurrences = 0,
    [datetime]$EndByDate,

    # Weekly options
    [switch]$Monday,
    [switch]$Tuesday,
    [switch]$Wednesday,
    [switch]$Thursday,
    [switch]$Friday,
    [switch]$Saturday,
    [switch]$Sunday,

    # Monthly options
    [int]$DayOfMonth = 0,
    [ValidateSet("", "First", "Second", "Third", "Fourth", "Last")]
    [string]$WeekNumber = "",
    [ValidateSet("", "Day", "Weekday", "WeekendDay", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")]
    [string]$DayOfWeek = ""
)

# Outlook Calendar Recurring Event Script
# Create a recurring calendar event
# Usage: .\outlook-calendar-recurring.ps1 -Subject "Daily standup" -Start "2026-02-10 09:00" -Pattern Daily
# Weekly: .\outlook-calendar-recurring.ps1 -Subject "Team meeting" -Start "2026-02-10 10:00" -Pattern Weekly -Monday -Wednesday -Friday
# Monthly: .\outlook-calendar-recurring.ps1 -Subject "Review" -Start "2026-02-10 14:00" -Pattern Monthly -DayOfMonth 15
# With end: .\outlook-calendar-recurring.ps1 -Subject "Sprint" -Start "2026-02-10" -Pattern Weekly -Occurrences 10

# Recurrence pattern constants
# olRecursDaily = 0
# olRecursWeekly = 1
# olRecursMonthly = 2
# olRecursMonthNth = 3 (e.g., "second Tuesday of month")
# olRecursYearly = 5
# olRecursYearNth = 6

# Day of week mask
# olSunday = 1, olMonday = 2, olTuesday = 4, olWednesday = 8
# olThursday = 16, olFriday = 32, olSaturday = 64

$dayMask = @{
    "Sunday" = 1
    "Monday" = 2
    "Tuesday" = 4
    "Wednesday" = 8
    "Thursday" = 16
    "Friday" = 32
    "Saturday" = 64
    "Day" = 127
    "Weekday" = 62
    "WeekendDay" = 65
}

$weekNumberMap = @{
    "First" = 1
    "Second" = 2
    "Third" = 3
    "Fourth" = 4
    "Last" = 5
}

$hasWeekNumber = -not [string]::IsNullOrWhiteSpace($WeekNumber)
$hasDayOfWeek = -not [string]::IsNullOrWhiteSpace($DayOfWeek)

if ($Pattern -eq "Monthly" -and ($hasWeekNumber -xor $hasDayOfWeek)) {
    throw "For monthly nth-weekday recurrence, provide both -WeekNumber and -DayOfWeek."
}

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
    $appointment.Duration = $Duration

    if ($Location) {
        $appointment.Location = $Location
    }

    if ($Body) {
        $appointment.Body = $Body
    }

    # Get recurrence pattern
    $recurrence = $appointment.GetRecurrencePattern()

    # Set the pattern type and options
    switch ($Pattern) {
        "Daily" {
            $recurrence.RecurrenceType = 0  # olRecursDaily
            $recurrence.Interval = $Interval
        }

        "Weekly" {
            $recurrence.RecurrenceType = 1  # olRecursWeekly
            $recurrence.Interval = $Interval

            # Calculate day of week mask
            $mask = 0
            if ($Monday) { $mask += 2 }
            if ($Tuesday) { $mask += 4 }
            if ($Wednesday) { $mask += 8 }
            if ($Thursday) { $mask += 16 }
            if ($Friday) { $mask += 32 }
            if ($Saturday) { $mask += 64 }
            if ($Sunday) { $mask += 1 }

            # If no days specified, use the day of the start date
            if ($mask -eq 0) {
                $startDayName = $Start.DayOfWeek.ToString()
                $mask = $dayMask[$startDayName]
            }

            $recurrence.DayOfWeekMask = $mask
        }

        "Monthly" {
            if ($hasWeekNumber -and $hasDayOfWeek) {
                # "Third Tuesday of month" style
                $recurrence.RecurrenceType = 3  # olRecursMonthNth
                $recurrence.Instance = $weekNumberMap[$WeekNumber]
                $recurrence.DayOfWeekMask = $dayMask[$DayOfWeek]
            } else {
                # "15th of each month" style
                $recurrence.RecurrenceType = 2  # olRecursMonthly
                if ($DayOfMonth -gt 0) {
                    $recurrence.DayOfMonth = $DayOfMonth
                } else {
                    $recurrence.DayOfMonth = $Start.Day
                }
            }
            $recurrence.Interval = $Interval
        }

        "Yearly" {
            if ($hasWeekNumber -and $hasDayOfWeek) {
                $recurrence.RecurrenceType = 6  # olRecursYearNth
                $recurrence.Instance = $weekNumberMap[$WeekNumber]
                $recurrence.DayOfWeekMask = $dayMask[$DayOfWeek]
                $recurrence.MonthOfYear = $Start.Month
            } else {
                $recurrence.RecurrenceType = 5  # olRecursYearly
                $recurrence.DayOfMonth = $Start.Day
                $recurrence.MonthOfYear = $Start.Month
            }
            $recurrence.Interval = $Interval
        }
    }

    # Set end condition
    if ($Occurrences -gt 0) {
        $recurrence.Occurrences = $Occurrences
    } elseif ($EndByDate) {
        $recurrence.PatternEndDate = $EndByDate
    } else {
        # Default: no end date (but Outlook may set a default)
        $recurrence.NoEndDate = $true
    }

    # Add attendees if provided (makes it a meeting)
    if ($Attendees) {
        $attendeeList = $Attendees -split "[;,]" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($attendee in $attendeeList) {
            $appointment.Recipients.Add($attendee)
        }
        $appointment.MeetingStatus = 1  # olMeeting
    }

    $appointment.Save()

    # If meeting, send invites
    if ($Attendees) {
        $appointment.Send()
    }

    Write-Host "`n=== RECURRING EVENT CREATED ===" -ForegroundColor Green
    Write-Host "Subject: $Subject" -ForegroundColor Yellow
    Write-Host "First occurrence: $($Start.ToString('ddd, MMM d, yyyy'))"
    Write-Host "Time: $($Start.ToString('h:mm tt')) ($Duration min)" -ForegroundColor Gray

    if ($Location) {
        Write-Host "Location: $Location" -ForegroundColor Gray
    }

    Write-Host "`nRecurrence:" -ForegroundColor Cyan

    switch ($Pattern) {
        "Daily" {
            if ($Interval -eq 1) {
                Write-Host "  Every day" -ForegroundColor Gray
            } else {
                Write-Host "  Every $Interval days" -ForegroundColor Gray
            }
        }

        "Weekly" {
            $days = @()
            if ($Monday -or ($mask -band 2)) { $days += "Mon" }
            if ($Tuesday -or ($mask -band 4)) { $days += "Tue" }
            if ($Wednesday -or ($mask -band 8)) { $days += "Wed" }
            if ($Thursday -or ($mask -band 16)) { $days += "Thu" }
            if ($Friday -or ($mask -band 32)) { $days += "Fri" }
            if ($Saturday -or ($mask -band 64)) { $days += "Sat" }
            if ($Sunday -or ($mask -band 1)) { $days += "Sun" }

            $daysList = $days -join ", "
            if ($Interval -eq 1) {
                Write-Host "  Every week on $daysList" -ForegroundColor Gray
            } else {
                Write-Host "  Every $Interval weeks on $daysList" -ForegroundColor Gray
            }
        }

        "Monthly" {
            if ($hasWeekNumber -and $hasDayOfWeek) {
                Write-Host "  $WeekNumber $DayOfWeek of every $Interval month(s)" -ForegroundColor Gray
            } else {
                $day = if ($DayOfMonth -gt 0) { $DayOfMonth } else { $Start.Day }
                Write-Host "  Day $day of every $Interval month(s)" -ForegroundColor Gray
            }
        }

        "Yearly" {
            Write-Host "  Every $Interval year(s) on $($Start.ToString('MMMM d'))" -ForegroundColor Gray
        }
    }

    # End condition
    if ($Occurrences -gt 0) {
        Write-Host "  Ends after $Occurrences occurrences" -ForegroundColor Gray
    } elseif ($EndByDate) {
        Write-Host "  Ends on $($EndByDate.ToString('MMM d, yyyy'))" -ForegroundColor Gray
    } else {
        Write-Host "  No end date" -ForegroundColor Gray
    }

    if ($Attendees) {
        Write-Host "`nMeeting invites sent to:" -ForegroundColor Cyan
        Write-Host "  $Attendees" -ForegroundColor Gray
    }
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
