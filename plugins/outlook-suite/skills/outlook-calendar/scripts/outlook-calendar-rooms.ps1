param(
    [Parameter(Mandatory=$true)]
    [datetime]$Start,

    [datetime]$End,
    [int]$Duration = 60,
    [string]$Building = "",
    [ValidateRange(0, [int]::MaxValue)]
    [int]$Capacity = 0,
    [int]$Limit = 200,
    [switch]$Json
)

# Outlook Calendar Rooms: which conference rooms are free for a time slot.
# Reads the organization's room list (Exchange "All Rooms") and each room's free/busy.
# Usage: .\outlook-calendar-rooms.ps1 -Start "2026-02-10 10:00"
# With end time: .\outlook-calendar-rooms.ps1 -Start "2026-02-10 10:00" -End "2026-02-10 11:30"
# Filter by building (name contains): .\outlook-calendar-rooms.ps1 -Start "2026-02-10 10:00" -Building "Main Hall"
# Book one: put its address in -Attendees of outlook-calendar-create.ps1.

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    if (-not $End) { $End = $Start.AddMinutes($Duration) }
    if ($End -le $Start) { throw "End time must be later than start time." }
    if ($Capacity -gt 0) { Write-Warning "Outlook's room list does not expose capacity; -Capacity $Capacity is ignored." }
    $null = Connect-Outlook

    $roomList = $null
    foreach ($l in $Namespace.AddressLists) {
        if ($l.Name -eq "All Rooms") { $roomList = $l; break }
        if (-not $roomList -and ($l.Name -like "*Room*" -or $l.Name -like "*Conference*")) { $roomList = $l }
    }
    if (-not $roomList) {
        Stop-WithError "No room list found in your organization's address book." "Rooms need Exchange room mailboxes in a room list. You can still invite a room by its email address."
    }

    # FreeBusy strings start at midnight of the start date, one character per $step minutes.
    $step = 15
    $first = [int][Math]::Floor(($Start - $Start.Date).TotalMinutes / $step)
    $last = [int][Math]::Ceiling(($End - $Start.Date).TotalMinutes / $step) - 1

    $rooms = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $roomList.AddressEntries) {
        if ($rooms.Count -ge $Limit) { break }
        $name = [string]$entry.Name
        if ($Building -and $name -notlike "*$Building*") { continue }
        $smtp = ""
        try { $u = $entry.GetExchangeUser(); if ($u) { $smtp = $u.PrimarySmtpAddress } } catch { }
        if (-not $smtp) { $smtp = [string]$entry.Address }
        $state = "unknown"
        try {
            $fb = [string]$entry.GetFreeBusy($Start.Date, $step, $true)
            if ($fb.Length -gt $last) {
                $slice = $fb.Substring($first, $last - $first + 1)
                $state = if ($slice -match "[23]") { "busy" } elseif ($slice -match "1") { "tentative" } else { "free" }
            }
        } catch { }
        $rooms.Add([ordered]@{ Name = $name; Email = $smtp; Status = $state })
    }

    if ($Json) {
        $isoStart = Format-MailDateIso $Start; $isoEnd = Format-MailDateIso $End
        $o = [ordered]@{ Start = $isoStart; End = $isoEnd; RoomList = [string]$roomList.Name }
        $o.Free = @($rooms | Where-Object { $_.Status -eq "free" }).Count
        $o.Rooms = $rooms.ToArray()
        Write-Json $o
        exit 0
    }

    Write-Host "`n=== ROOMS: $($Start.ToString('ddd, MMM d h:mm tt')) - $($End.ToString('h:mm tt')) ===" -ForegroundColor Cyan
    Write-Host "Room list: $($roomList.Name)$(if ($Building) { "  (name contains '$Building')" })" -ForegroundColor Gray
    if ($rooms.Count -eq 0) { Write-Host "No rooms match." -ForegroundColor Yellow; exit 0 }
    foreach ($s in "free", "tentative", "busy", "unknown") {
        $group = @($rooms | Where-Object { $_.Status -eq $s })
        if ($group.Count -eq 0) { continue }
        $color = @{ "free" = "Green"; "tentative" = "Yellow"; "busy" = "Red"; "unknown" = "DarkGray" }[$s]
        Write-Host "`n--- $($s.Substring(0,1).ToUpper() + $s.Substring(1)) ($($group.Count)) ---" -ForegroundColor $color
        foreach ($r in $group) { Write-Host "  $($r.Name)  <$($r.Email)>" -ForegroundColor $color }
    }
    Write-Host "`nTo book a room, add its address to -Attendees in outlook-calendar-create.ps1." -ForegroundColor Gray
}
catch {
    Stop-WithError $_
}
