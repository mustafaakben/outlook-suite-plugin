param(
    [Parameter(Mandatory=$true)]
    [string]$EntryID,

    [Parameter(Mandatory=$true)]
    [string]$Attendees,

    [string]$Optional = "",
    [string]$Subject = "",
    [string]$Location = "",
    [string]$Body = "",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Calendar Hold to Invite: turn one of the user's holds into a meeting and send the
# invitations. Sends mail, so it previews (and changes nothing) unless -Confirm.
# For a series hold, the whole series becomes a meeting series.
# Usage: .\outlook-calendar-hold-to-invite.ps1 -EntryID "00000000..." -Attendees "Jane Doe; john@example.com"
# Send: .\outlook-calendar-hold-to-invite.ps1 -EntryID "00000000..." -Attendees "jane@example.com" -Subject "Project sync" -Location "Zoom" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    $item = Get-CalendarItem -Namespace $Namespace -EntryID $EntryID
    $rec = ConvertTo-EventRecord $item
    if ($rec.Kind -ne "hold") { Stop-WithError "This event is a $($rec.Kind), not a hold. Only a hold with no one else invited can be turned into an invitation." }

    $bad = @(); $people = @()
    foreach ($pair in @(@(Split-AttendeeList $Attendees | ForEach-Object { @{ T = $_; Type = 1 } }) + @(Split-AttendeeList $Optional | ForEach-Object { @{ T = $_; Type = 2 } }))) {
        $a = Resolve-CalendarAttendee -Namespace $Namespace -Text $pair.T
        if ($a.Resolved) { $people += [ordered]@{ Name = $a.Name; Email = $a.Email; Type = $pair.Type } } else { $bad += "$($pair.T): $($a.Note)" }
    }
    if ($bad.Count) { Stop-WithError "Some attendees could not be resolved. Nothing was changed or sent.`n  $($bad -join "`n  ")" }
    if ($people.Count -eq 0) { Stop-WithError "Give at least one attendee." }

    $invite = ($people | ForEach-Object { "$($_.Name) <$($_.Email)>" + $(if ($_.Type -eq 2) { " (optional)" } else { "" }) }) -join "; "
    $details = [ordered]@{ "Would invite" = $invite }
    if ($Subject) { $details.Subject = $Subject }
    if ($Location) { $details.Location = $Location }
    if ($Body) { $details.Body = "REPLACES the hold's notes: " + $(if ($Body.Length -gt 120) { $Body.Substring(0, 120) + "..." } else { $Body }) }
    if ($rec.Recurring) { $details.Scope = "the whole series" }
    if ((Confirm-CalendarAction -Action "send" -Title "TURN HOLD INTO MEETING" -Item $item -Confirmed $Confirm -Details $details -What "the invitations" -Json:$Json) -ne $true) { Write-CalendarPreview; exit 0 }

    if ($Subject) { $item.Subject = $Subject }
    if ($Location) { $item.Location = $Location }
    if ($Body) { $item.Body = $Body }
    $item.MeetingStatus = 1   # olMeeting
    foreach ($p in $people) { $r = $item.Recipients.Add($p.Email); $r.Type = $p.Type }
    Assert-OutlookRecipients $item
    $item.Save()
    $item.Send()

    $rec = ConvertTo-EventRecord $item
    if ($Json) { $o = ConvertTo-EventJson $rec; $o.Action = "hold-to-invite"; $o.Result = "invites sent"; $o.Invited = @($people | ForEach-Object { $_.Email }); Write-Json $o; exit 0 }
    Write-Host "=== HOLD TURNED INTO MEETING, INVITES SENT ===" -ForegroundColor Green
    Write-EventRecord $rec -Prefix "  "
    Write-Host "      Invited: $invite" -ForegroundColor Cyan
}
catch {
    Stop-WithError $_
}
