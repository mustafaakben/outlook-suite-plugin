param(
    [Parameter(Mandatory=$true)]
    [string]$EntryID,
    [datetime]$OccurrenceStart,
    [string]$Message = "",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Calendar Cancel: cancel a meeting the user organizes and notify every attendee
# (Outlook's "Cancel meeting": MeetingStatus = olMeetingCanceled + Send), with an optional
# message at the top of the notice. Sends mail, so it previews (and changes nothing)
# unless -Confirm. A series EntryID cancels the whole series; add -OccurrenceStart to
# cancel one occurrence.
# Not for invitations from others (use outlook-calendar-respond.ps1 -Response Decline) or
# holds (use outlook-calendar-delete.ps1).
# Usage: .\outlook-calendar-cancel.ps1 -EntryID "00000000..." -Message "Sorry, I need to cancel this week."
# Usage: .\outlook-calendar-cancel.ps1 -EntryID "00000000..." -OccurrenceStart "2026-11-03 09:00" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    $occ = $null; if ($PSBoundParameters.ContainsKey("OccurrenceStart")) { $occ = $OccurrenceStart }
    $item = Get-CalendarItem -Namespace $Namespace -EntryID $EntryID -OccurrenceStart $occ
    $rec = ConvertTo-EventRecord $item

    # Validate before touching the item.
    $ms = [int]$item.MeetingStatus
    if ($ms -eq 5 -or $ms -eq 7) { Stop-WithError "This meeting is already cancelled." }
    if ($ms -eq 3) { Stop-WithError "You were invited to this meeting; only the organizer can cancel it." "Decline instead: outlook-calendar-respond.ps1 -EntryID ... -Response Decline" }
    if ($rec.Kind -ne "meeting") { Stop-WithError "This is a $($rec.Kind) with no one else invited; there is no one to notify." "Remove it with outlook-calendar-delete.ps1." }

    $names = @($item.Recipients | ForEach-Object { [string]$_.Name } | Where-Object { $_ -and $_ -ne [string]$Namespace.CurrentUser.Name })
    $scope = "this meeting"
    if ($rec.Occurrence) { $scope = "this occurrence only" } elseif ($rec.Recurring) { $scope = "the ENTIRE SERIES (every occurrence)" }
    $details = [ordered]@{ Cancels = $scope; Notifies = "$($names.Count): $(($names | Select-Object -First 8) -join '; ')$(if ($names.Count -gt 8) { '; ...' })" }
    if ($Message) { $details.Message = $Message }
    if ((Confirm-CalendarAction -Action "send" -Title "CANCEL MEETING" -Item $item -Confirmed $Confirm -Details $details -What "the cancellation notice" -Json:$Json) -ne $true) { Write-CalendarPreview; exit 0 }

    if ($Message) { $item.Body = $Message + "`r`n`r`n" + [string]$item.Body }
    $item.MeetingStatus = 5   # olMeetingCanceled
    $item.Send()
    try { $item.Delete() } catch { }   # remove the organizer's cancelled copy, as Outlook does

    if ($Json) { $o = ConvertTo-EventJson $rec; $o.Action = "cancel"; $o.Result = "cancelled, attendees notified"; $o.Scope = $scope; $o.Notified = $names.Count; Write-Json $o; exit 0 }
    Write-Host "=== MEETING CANCELLED ($scope), $($names.Count) attendees notified ===" -ForegroundColor Green
    Write-EventRecord $rec -Prefix "  "
}
catch {
    if ($item) { try { $item.Close(1) } catch { } }   # never leave unsaved edits on the COM item
    Stop-WithError $_
}
