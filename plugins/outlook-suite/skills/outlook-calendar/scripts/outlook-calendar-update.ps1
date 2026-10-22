param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [datetime]$OccurrenceStart,

    [int]$Days = 7,
    [string]$Subject = "",
    [datetime]$Start,
    [datetime]$End,
    [string]$Location = "",
    [string]$Body = "",
    [ValidateSet("", "Free", "Tentative", "Busy", "OutOfOffice", "WorkingElsewhere")]
    [string]$BusyStatus = "",
    [string]$Categories = "",
    [int]$Reminder = -1,
    [switch]$NoReminder,
    [switch]$SendUpdate,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Calendar Update: change an event's subject, time, location, body, show-as,
# categories or reminder. Saved locally; for a meeting you organize, -SendUpdate also
# sends the update to attendees (previews unless -Confirm).
# Usage: .\outlook-calendar-update.ps1 -EntryID "00000000..." -Subject "New Title"
# Move (keeps the duration): .\outlook-calendar-update.ps1 -EntryID "00000000..." -Start "2026-02-05 14:00"
# One occurrence of a series: .\outlook-calendar-update.ps1 -EntryID "00000000..." -OccurrenceStart "2026-02-05 09:00" -Location "Room 101"
# Meeting, tell attendees: .\outlook-calendar-update.ps1 -EntryID "00000000..." -Start "2026-02-05 15:00" -SendUpdate -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    $occ = $null; if ($PSBoundParameters.ContainsKey("OccurrenceStart")) { $occ = $OccurrenceStart }
    $item = Get-CalendarItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Days $Days -OccurrenceStart $occ

    # Validate BEFORE touching the item: an early exit after in-memory edits would leave a
    # dirty COM item that the warm host could save later.
    $wanted = $Subject -or $Location -or $Body -or $BusyStatus -or $Categories -or $NoReminder -or ($Reminder -ge 0) -or $PSBoundParameters.ContainsKey("Start") -or $PSBoundParameters.ContainsKey("End")
    if (-not $wanted) { Stop-WithError "No changes given." "Use -Subject, -Start, -End, -Location, -Body, -BusyStatus, -Categories, -Reminder or -NoReminder." }
    $others = Get-OtherAttendeeCount -Names @($item.Recipients | ForEach-Object { [string]$_.Name }) -Organizer ([string]$item.Organizer)
    $isOrganizer = ([int]$item.MeetingStatus -eq 1 -and $others -gt 0)
    if ($SendUpdate -and -not $isOrganizer) { Stop-WithError "-SendUpdate works only for a meeting you organize that has attendees. Nothing was changed." }

    $changes = [ordered]@{}
    if ($Subject) { $changes.Subject = "'$($item.Subject)' -> '$Subject'"; $item.Subject = $Subject }
    if ($PSBoundParameters.ContainsKey("Start")) {
        $oldStart = [datetime]$item.Start
        $duration = ([datetime]$item.End) - $oldStart
        $item.Start = $Start
        if (-not $PSBoundParameters.ContainsKey("End")) { $item.End = $Start.Add($duration) }
        $changes.Start = "$($oldStart.ToString('ddd MMM d h:mm tt')) -> $($Start.ToString('ddd MMM d h:mm tt'))"
    }
    if ($PSBoundParameters.ContainsKey("End")) {
        $changes.End = "$(([datetime]$item.End).ToString('h:mm tt')) -> $($End.ToString('h:mm tt'))"
        $item.End = $End
    }
    if ($Location) { $changes.Location = "'$($item.Location)' -> '$Location'"; $item.Location = $Location }
    if ($Body) { $item.Body = $Body; $changes.Body = "updated" }
    if ($BusyStatus) { $changes.ShowAs = "$($script:BusyNames[[int]$item.BusyStatus]) -> $BusyStatus"; $item.BusyStatus = $script:BusyValues[$BusyStatus] }
    if ($Categories) { $changes.Categories = "'$($item.Categories)' -> '$Categories'"; $item.Categories = $Categories }
    if ($NoReminder) { $item.ReminderSet = $false; $changes.Reminder = "off" }
    elseif ($Reminder -ge 0) { $item.ReminderSet = $true; $item.ReminderMinutesBeforeStart = $Reminder; $changes.Reminder = "$Reminder minutes before" }

    $sent = $false
    if ($SendUpdate) {
        if ((Confirm-CalendarAction -Action "send" -Title "SEND MEETING UPDATE" -Item $item -Confirmed $Confirm -Details $changes -What "the update to attendees" -Json:$Json) -ne $true) { $item.Close(1); Write-CalendarPreview; exit 0 }
        $item.Save()
        try { $item.Send() }
        catch { Stop-WithError "Saved in your calendar, but the update was NOT sent to attendees: $($_.Exception.Message)" "Open the meeting in Outlook and choose Send Update, or rerun with -SendUpdate -Confirm." }
        $sent = $true
    } else {
        $item.Save()
    }

    $rec = ConvertTo-EventRecord $item
    if ($Json) {
        $o = ConvertTo-EventJson $rec; $o.Action = "update"; $o.Result = $(if ($sent) { "updated, sent to attendees" } else { "updated" }); $o.Changes = $changes
        Write-Json $o; exit 0
    }
    Write-Host "=== EVENT UPDATED ===" -ForegroundColor Green
    Write-EventRecord $rec -Prefix "  " -Detail
    Write-Host "  Changes:" -ForegroundColor Cyan
    foreach ($k in $changes.Keys) { Write-Host "    ${k}: $($changes[$k])" -ForegroundColor Gray }
    if ($sent) { Write-Host "  Update sent to attendees." -ForegroundColor Cyan }
    elseif ($isOrganizer) { Write-Host "  Saved in your calendar only. Attendees were not told; add -SendUpdate -Confirm to send the update." -ForegroundColor Yellow }
}
catch {
    if ($item) { try { $item.Close(1) } catch { } }   # never leave unsaved edits on the COM item
    Stop-WithError $_
}
