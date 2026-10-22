param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [datetime]$OccurrenceStart,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Accept", "Tentative", "Decline")]
    [string]$Response,

    [int]$Days = 7,
    [switch]$NoResponse,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Calendar Respond: accept, tentatively accept or decline a meeting invitation.
# Sends the response to the organizer. Irreversible: previews unless -Confirm.
# -NoResponse changes only your copy and never calls Outlook's Respond(): Accept/Tentative
# set the event's show-as, Decline removes it from your calendar.
# Usage: .\outlook-calendar-respond.ps1 -EntryID "00000000..." -Response Accept -Confirm
# Decline one occurrence: .\outlook-calendar-respond.ps1 -EntryID "00000000..." -OccurrenceStart "2026-02-05 09:00" -Response Decline -Confirm
# Without telling the organizer: .\outlook-calendar-respond.ps1 -EntryID "00000000..." -Response Tentative -NoResponse -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

$codes = @{ "Accept" = 3; "Tentative" = 2; "Decline" = 4 }   # olMeetingAccepted / Tentative / Declined

try {
    $null = Connect-Outlook
    $occ = $null; if ($PSBoundParameters.ContainsKey("OccurrenceStart")) { $occ = $OccurrenceStart }
    $item = Get-CalendarItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Days $Days -OccurrenceStart $occ
    $ms = [int]$item.MeetingStatus
    if ($ms -eq 0) { Stop-WithError "This is an appointment, not a meeting invitation. Nothing to respond to." }
    if ($ms -eq 1) { Stop-WithError "You organize this meeting. Use outlook-calendar-update.ps1 or outlook-calendar-delete.ps1 instead." }

    $send = -not $NoResponse
    $what = if ($send) { "the '$Response' response to the organizer" } else { "the '$Response' response (not sent)" }
    $details = [ordered]@{ Response = $Response; "Tell organizer" = $(if ($send) { "yes" } else { "no" }) }
    $rec = ConvertTo-EventRecord $item
    if ($rec.Recurring -and -not $rec.Occurrence) { $details.Scope = "the ENTIRE SERIES (every occurrence); add -OccurrenceStart for one" }
    # -NoResponse: the show-as changes only when it does not already say what you mean
    # (Accept: Free/Tentative -> Busy; Tentative: Free/Busy -> Tentative; Out of Office and
    # Working Elsewhere are never touched).
    $oldBusy = [int]$item.BusyStatus; $newBusy = $oldBusy
    if (-not $send -and $Response -eq "Accept" -and $oldBusy -in 0, 1) { $newBusy = 2 }
    if (-not $send -and $Response -eq "Tentative" -and $oldBusy -in 0, 2) { $newBusy = 1 }
    if (-not $send -and $Response -ne "Decline") { $details.ShowAs = "$($script:BusyNames[$oldBusy]) -> $($script:BusyNames[$newBusy])" }
    if ((Confirm-CalendarAction -Action "respond" -Title "RESPOND TO MEETING" -Item $item -Confirmed $Confirm -Details $details -What $what -Json:$Json) -ne $true) { Write-CalendarPreview; exit 0 }
    if ($send) {
        # Respond(code, fNoUI = $true) returns the response item; Send() sends it.
        $reply = $item.Respond($codes[$Response], $true)
        if ($reply) { $reply.Send() }
    } else {
        # -NoResponse never calls Respond(): Microsoft's docs disagree on whether
        # Respond(code, $true) sends by itself. Only the local copy changes instead:
        # Accept / Tentative adjust show-as (see above) and save; Decline removes the event.
        if ($Response -eq "Decline") { $item.Delete() }
        elseif ($newBusy -ne $oldBusy) { $item.BusyStatus = $newBusy; $item.Save() }
    }

    if ($Json) {
        $o = ConvertTo-EventJson $rec; $o.Action = "respond"; $o.Response = $Response; $o.Sent = $send
        if ($details.Contains("Scope")) { $o.Scope = $details.Scope }
        if ($details.Contains("ShowAs")) { $o.ShowAs = $details.ShowAs }
        Write-Json $o; exit 0
    }
    $color = @{ "Accept" = "Green"; "Tentative" = "Yellow"; "Decline" = "Red" }[$Response]
    Write-Host "=== MEETING RESPONSE: $($Response.ToUpper()) ===" -ForegroundColor $color
    Write-EventRecord $rec -Prefix "  "
    if ($details.Contains("Scope")) { Write-Host "  Scope: $($details.Scope)" -ForegroundColor Cyan }
    if ($details.Contains("ShowAs")) { Write-Host "  Shows as: $($details.ShowAs)" -ForegroundColor Cyan }
    if ($send) { Write-Host "  Response sent to the organizer." -ForegroundColor Cyan } else { Write-Host "  Changed in your calendar only; the organizer was not told." -ForegroundColor Gray }
}
catch {
    Stop-WithError $_
}
