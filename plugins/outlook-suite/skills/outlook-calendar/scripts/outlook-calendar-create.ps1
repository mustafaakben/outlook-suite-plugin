param(
    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [Parameter(Mandatory=$true)]
    [datetime]$Start,

    [datetime]$End,
    [int]$Duration = 0,
    [string]$Location = "",
    [string]$Body = "",
    [switch]$AllDay,
    [int]$Reminder = 15,
    [switch]$NoReminder,
    [ValidateSet("", "Free", "Tentative", "Busy", "OutOfOffice", "WorkingElsewhere")]
    [string]$BusyStatus = "",
    [string]$Categories = "",
    [string]$Attendees = "",
    [string]$Optional = "",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Calendar Create: an appointment (a "hold": no attendees) or a meeting invite.
# A hold is saved right away. A meeting sends invitations, so it previews unless -Confirm.
# Usage: .\outlook-calendar-create.ps1 -Subject "Focus block" -Start "2026-02-05 10:00" -Duration 90
# With end time: .\outlook-calendar-create.ps1 -Subject "Meeting" -Start "2026-02-05 10:00" -End "2026-02-05 11:00"
# All-day event: .\outlook-calendar-create.ps1 -Subject "Vacation" -Start "2026-02-10" -AllDay -BusyStatus OutOfOffice
# Tentative hold, no reminder: .\outlook-calendar-create.ps1 -Subject "Backup slot" -Start "2026-02-05 17:00" -Duration 30 -BusyStatus Tentative -NoReminder
# Meeting invite: .\outlook-calendar-create.ps1 -Subject "Team Meeting" -Start "2026-02-05 14:00" -Attendees "john@example.com; Jane Doe" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    $appt = $Outlook.CreateItem(1)   # olAppointmentItem
    $appt.Subject = $Subject
    $appt.Start = $Start
    if ($End) { $appt.End = $End }
    elseif ($AllDay) { $appt.End = $Start.Date.AddDays(1) }
    elseif ($Duration -gt 0) { $appt.End = $Start.AddMinutes($Duration) }
    else { $appt.End = $Start.AddHours(1) }
    if ($AllDay) { $appt.AllDayEvent = $true }
    if ($Location) { $appt.Location = $Location }
    if ($Body) { $appt.Body = $Body }
    if ($BusyStatus) { $appt.BusyStatus = $script:BusyValues[$BusyStatus] }
    if ($Categories) { $appt.Categories = $Categories }
    if ($NoReminder) { $appt.ReminderSet = $false }
    else { $appt.ReminderSet = $true; $appt.ReminderMinutesBeforeStart = $Reminder }

    $names = @(Split-AttendeeList $Attendees)
    $optionalNames = @(Split-AttendeeList $Optional)
    if ($names.Count -eq 0 -and $optionalNames.Count -eq 0) {
        $appt.Save()
        $rec = ConvertTo-EventRecord $appt
        if ($Json) { $o = ConvertTo-EventJson $rec; $o.Action = "create"; $o.Result = "created"; Write-Json $o; exit 0 }
        Write-Host "=== APPOINTMENT CREATED ===" -ForegroundColor Green
        Write-EventRecord $rec -Prefix "  " -Detail
        if (-not $NoReminder) { Write-Host "      Reminder: $Reminder minutes before" -ForegroundColor Gray }
        exit 0
    }

    # Meeting: resolve every attendee before anything is saved or sent.
    $bad = @()
    $resolved = @()
    foreach ($pair in @(@($names | ForEach-Object { @{ T = $_; Type = 1 } }) + @($optionalNames | ForEach-Object { @{ T = $_; Type = 2 } }))) {
        $a = Resolve-CalendarAttendee -Namespace $Namespace -Text $pair.T
        if (-not $a.Resolved) { $bad += "$($pair.T): $($a.Note)"; continue }
        $resolved += [ordered]@{ Name = $a.Name; Email = $a.Email; Type = $pair.Type }
    }
    if ($bad.Count -gt 0) { Stop-WithError "Some attendees could not be resolved. Nothing was saved or sent.`n  $($bad -join "`n  ")" }

    $appt.MeetingStatus = 1   # olMeeting
    foreach ($a in $resolved) { $r = $appt.Recipients.Add($a.Email); $r.Type = $a.Type }
    $list = ($resolved | ForEach-Object { "$($_.Name) <$($_.Email)>" + $(if ($_.Type -eq 2) { " (optional)" } else { "" }) }) -join "; "
    if (-not (Test-ConfirmRequired -Action "send") -or $Confirm) {
        Assert-OutlookRecipients $appt
        $appt.Save()
        $appt.Send()
        $rec = ConvertTo-EventRecord $appt
        if ($Json) { $o = ConvertTo-EventJson $rec; $o.Action = "create"; $o.Result = "invites sent"; $o.Invited = @($resolved | ForEach-Object { $_.Email }); Write-Json $o; exit 0 }
        Write-Host "=== MEETING CREATED, INVITES SENT ===" -ForegroundColor Green
        Write-EventRecord $rec -Prefix "  " -Detail
        Write-Host "      Invited: $list" -ForegroundColor Cyan
        exit 0
    }
    # Preview: nothing saved, nothing sent.
    if ($Json) {
        Set-CalendarPreviewJson -Action "send" -What "creating the meeting and sending invitations" -Rec (ConvertTo-EventRecord $appt) -Details ([ordered]@{ WouldInvite = @($resolved | ForEach-Object { $_.Email }) })
        $appt.Close(1); Write-CalendarPreview; exit 0
    }
    Write-Host "=== NEW MEETING (PREVIEW) ===" -ForegroundColor Yellow
    Write-EventRecord (ConvertTo-EventRecord $appt) -Prefix "  "
    Write-Host "      Would invite: $list" -ForegroundColor Cyan
    $null = Approve-Action -Action "send" -Confirmed $false -What "creating the meeting and sending invitations"
    $appt.Close(1)   # olDiscard
}
catch {
    Stop-WithError $_
}
