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
    [ValidateSet("", "Free", "Tentative", "Busy", "OutOfOffice", "WorkingElsewhere")]
    [string]$BusyStatus = "",
    [string]$Categories = "",
    [switch]$NoReminder,
    [switch]$Confirm,
    [switch]$Json,

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

# Outlook Calendar Recurring: create a recurring hold or meeting series.
# Without -Attendees it is saved right away. With -Attendees it sends invitations,
# so it previews unless -Confirm.
# Usage: .\outlook-calendar-recurring.ps1 -Subject "Daily standup" -Start "2026-02-10 09:00" -Pattern Daily
# Weekly: .\outlook-calendar-recurring.ps1 -Subject "Team meeting" -Start "2026-02-10 10:00" -Pattern Weekly -Monday -Wednesday -Friday
# Monthly: .\outlook-calendar-recurring.ps1 -Subject "Review" -Start "2026-02-10 14:00" -Pattern Monthly -DayOfMonth 15
# Nth weekday: .\outlook-calendar-recurring.ps1 -Subject "Board" -Start "2026-02-10 14:00" -Pattern Monthly -WeekNumber Second -DayOfWeek Tuesday
# With end: .\outlook-calendar-recurring.ps1 -Subject "Sprint" -Start "2026-02-10" -Pattern Weekly -Occurrences 10
# Meeting series: .\outlook-calendar-recurring.ps1 -Subject "1:1" -Start "2026-02-10 10:00" -Pattern Weekly -Attendees "jane@example.com" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

# Day of week mask: olSunday = 1, olMonday = 2, ... olSaturday = 64
$dayMask = @{
    "Sunday" = 1; "Monday" = 2; "Tuesday" = 4; "Wednesday" = 8; "Thursday" = 16; "Friday" = 32; "Saturday" = 64
    "Day" = 127; "Weekday" = 62; "WeekendDay" = 65
}
$weekNumberMap = @{ "First" = 1; "Second" = 2; "Third" = 3; "Fourth" = 4; "Last" = 5 }
$hasWeekNumber = -not [string]::IsNullOrWhiteSpace($WeekNumber)
$hasDayOfWeek = -not [string]::IsNullOrWhiteSpace($DayOfWeek)

try {
    if ($Pattern -in @("Monthly", "Yearly") -and ($hasWeekNumber -xor $hasDayOfWeek)) {
        throw "For an nth-weekday recurrence, give both -WeekNumber and -DayOfWeek."
    }
    $null = Connect-Outlook
    $appt = $Outlook.CreateItem(1)   # olAppointmentItem
    $appt.Subject = $Subject
    $appt.Start = $Start
    $appt.Duration = $Duration
    if ($Location) { $appt.Location = $Location }
    if ($Body) { $appt.Body = $Body }
    if ($BusyStatus) { $appt.BusyStatus = $script:BusyValues[$BusyStatus] }
    if ($Categories) { $appt.Categories = $Categories }
    if ($NoReminder) { $appt.ReminderSet = $false }

    $rp = $appt.GetRecurrencePattern()
    $desc = ""
    $mask = 0
    switch ($Pattern) {
        "Daily" {
            $rp.RecurrenceType = 0   # olRecursDaily
            $rp.Interval = $Interval
            $desc = if ($Interval -eq 1) { "Every day" } else { "Every $Interval days" }
        }
        "Weekly" {
            $rp.RecurrenceType = 1   # olRecursWeekly
            $rp.Interval = $Interval
            if ($Monday) { $mask += 2 }; if ($Tuesday) { $mask += 4 }; if ($Wednesday) { $mask += 8 }
            if ($Thursday) { $mask += 16 }; if ($Friday) { $mask += 32 }; if ($Saturday) { $mask += 64 }; if ($Sunday) { $mask += 1 }
            if ($mask -eq 0) { $mask = $dayMask[$Start.DayOfWeek.ToString()] }
            $rp.DayOfWeekMask = $mask
            $names = @(foreach ($d in "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday") { if ($mask -band $dayMask[$d]) { $d.Substring(0, 3) } })
            $every = if ($Interval -eq 1) { "Every week" } else { "Every $Interval weeks" }
            $desc = "$every on $($names -join ', ')"
        }
        "Monthly" {
            if ($hasWeekNumber) {
                $rp.RecurrenceType = 3   # olRecursMonthNth
                $rp.Instance = $weekNumberMap[$WeekNumber]
                $rp.DayOfWeekMask = $dayMask[$DayOfWeek]
                $desc = "$WeekNumber $DayOfWeek of every $Interval month(s)"
            } else {
                $rp.RecurrenceType = 2   # olRecursMonthly
                $day = if ($DayOfMonth -gt 0) { $DayOfMonth } else { $Start.Day }
                $rp.DayOfMonth = $day
                $desc = "Day $day of every $Interval month(s)"
            }
            $rp.Interval = $Interval
        }
        "Yearly" {
            if ($hasWeekNumber) {
                $rp.RecurrenceType = 6   # olRecursYearNth
                $rp.Instance = $weekNumberMap[$WeekNumber]
                $rp.DayOfWeekMask = $dayMask[$DayOfWeek]
                $rp.MonthOfYear = $Start.Month
                $desc = "$WeekNumber $DayOfWeek of $($Start.ToString('MMMM')) every $Interval year(s)"
            } else {
                $rp.RecurrenceType = 5   # olRecursYearly
                $rp.DayOfMonth = $Start.Day
                $rp.MonthOfYear = $Start.Month
                $desc = "Every $Interval year(s) on $($Start.ToString('MMMM d'))"
            }
            $rp.Interval = $Interval
        }
    }
    if ($Occurrences -gt 0) { $rp.Occurrences = $Occurrences; $desc += ", $Occurrences times" }
    elseif ($PSBoundParameters.ContainsKey("EndByDate")) { $rp.PatternEndDate = $EndByDate; $desc += ", until $($EndByDate.ToString('MMM d, yyyy'))" }
    else { $rp.NoEndDate = $true; $desc += ", no end date" }

    $names = @(Split-AttendeeList $Attendees)
    $resolved = @()
    if ($names.Count -gt 0) {
        $bad = @()
        foreach ($n in $names) {
            $a = Resolve-CalendarAttendee -Namespace $Namespace -Text $n
            if ($a.Resolved) { $resolved += $a } else { $bad += "${n}: $($a.Note)" }
        }
        if ($bad.Count -gt 0) { Stop-WithError "Some attendees could not be resolved. Nothing was saved or sent.`n  $($bad -join "`n  ")" }
        $appt.MeetingStatus = 1   # olMeeting
        foreach ($a in $resolved) { [void]$appt.Recipients.Add($a.Email) }
        if ($Confirm -or -not (Test-ConfirmRequired -Action "send")) {
            Assert-OutlookRecipients $appt
            $appt.Save()
            $appt.Send()
        } elseif ($Json) {
            Set-CalendarPreviewJson -Action "send" -What "creating the series and sending invitations" -Rec (ConvertTo-EventRecord $appt) -Details ([ordered]@{ Repeats = $desc; WouldInvite = @($resolved | ForEach-Object { $_.Email }) })
            $appt.Close(1)
            Write-CalendarPreview
            exit 0
        } else {
            Write-Host "=== NEW MEETING SERIES (PREVIEW) ===" -ForegroundColor Yellow
            Write-EventRecord (ConvertTo-EventRecord $appt) -Prefix "  "
            Write-Host "      Repeats: $desc" -ForegroundColor Cyan
            Write-Host "      Would invite: $(($resolved | ForEach-Object { "$($_.Name) <$($_.Email)>" }) -join '; ')" -ForegroundColor Cyan
            $null = Approve-Action -Action "send" -Confirmed $false -What "creating the series and sending invitations"
            $appt.Close(1)
            exit 0
        }
    } else {
        $appt.Save()
    }

    $rec = ConvertTo-EventRecord $appt
    if ($Json) {
        $o = ConvertTo-EventJson $rec; $o.Action = "recurring"; $o.Result = $(if ($resolved.Count) { "invites sent" } else { "created" }); $o.Repeats = $desc
        $o.Invited = @($resolved | ForEach-Object { $_.Email })
        Write-Json $o; exit 0
    }
    $title = if ($resolved.Count) { "RECURRING MEETING CREATED, INVITES SENT" } else { "RECURRING EVENT CREATED" }
    Write-Host "=== $title ===" -ForegroundColor Green
    Write-EventRecord $rec -Prefix "  "
    Write-Host "      Repeats: $desc" -ForegroundColor Cyan
    if ($resolved.Count) { Write-Host "      Invited: $(($resolved | ForEach-Object { $_.Email }) -join '; ')" -ForegroundColor Cyan }
}
catch {
    Stop-WithError $_
}
