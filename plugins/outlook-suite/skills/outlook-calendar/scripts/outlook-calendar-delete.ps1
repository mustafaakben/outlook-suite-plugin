param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [datetime]$OccurrenceStart,
    [int]$Days = 7,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Calendar Delete: remove an event (moves it to Deleted Items).
# Irreversible action: previews unless -Confirm. A series EntryID deletes the WHOLE series;
# add -OccurrenceStart to delete one occurrence only.
# Deleting a meeting you organize does not notify attendees; deleting one you were invited
# to does not send a decline. Use outlook-calendar-respond.ps1 -Response Decline for that.
# Usage: .\outlook-calendar-delete.ps1 -EntryID "00000000..." -Confirm
# One occurrence: .\outlook-calendar-delete.ps1 -EntryID "00000000..." -OccurrenceStart "2026-02-05 09:00" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    $occ = $null; if ($PSBoundParameters.ContainsKey("OccurrenceStart")) { $occ = $OccurrenceStart }
    $item = Get-CalendarItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Days $Days -OccurrenceStart $occ
    $rec = ConvertTo-EventRecord $item

    $scope = "this event"
    if ($rec.Occurrence) { $scope = "this occurrence only" }
    elseif ($rec.Recurring) { $scope = "the ENTIRE SERIES (every occurrence)" }
    $details = [ordered]@{ Deletes = $scope }
    if ($rec.Kind -eq "meeting") { $details.Note = "Attendees and the organizer are not notified." }
    if ((Confirm-CalendarAction -Action "delete" -Title "DELETE EVENT" -Item $item -Confirmed $Confirm -Details $details -What "the delete" -Json:$Json) -ne $true) { Write-CalendarPreview; exit 0 }

    $item.Delete()
    if ($Json) { $o = ConvertTo-EventJson $rec; $o.Action = "delete"; $o.Result = "deleted"; $o.Scope = $scope; Write-Json $o; exit 0 }
    Write-Host "=== EVENT DELETED ($scope) ===" -ForegroundColor Green
    Write-EventRecord $rec -Prefix "  "
}
catch {
    Stop-WithError $_
}
