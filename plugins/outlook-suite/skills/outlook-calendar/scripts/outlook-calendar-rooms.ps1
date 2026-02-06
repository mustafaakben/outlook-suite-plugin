param(
    [Parameter(Mandatory=$true)]
    [datetime]$Start,

    [datetime]$End,
    [string]$Building = "",
    [ValidateRange(0, [int]::MaxValue)]
    [int]$Capacity = 0
)

# Outlook Calendar Rooms Script
# Find available conference rooms
# Usage: .\outlook-calendar-rooms.ps1 -Start "2026-02-10 10:00"
# With end time: .\outlook-calendar-rooms.ps1 -Start "2026-02-10 10:00" -End "2026-02-10 11:00"
# Filter by building: .\outlook-calendar-rooms.ps1 -Start "2026-02-10 10:00" -Building "HQ"

# Note: This feature requires Exchange/Office 365 with room mailboxes configured

$Outlook = $null
$Namespace = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")

    # Default end time is 1 hour after start
    if (-not $End) {
        $End = $Start.AddHours(1)
    }

    if ($End -le $Start) {
        throw "End time must be later than start time."
    }

    $freeBusyIntervalMinutes = 30
    if ($freeBusyIntervalMinutes -lt 1) {
        throw "Free/busy interval must be at least 1 minute."
    }

    Write-Host "`n=== FIND AVAILABLE ROOMS ===" -ForegroundColor Cyan
    Write-Host "Date: $($Start.ToString('ddd, MMM d, yyyy'))"
    Write-Host "Time: $($Start.ToString('h:mm tt')) - $($End.ToString('h:mm tt'))" -ForegroundColor Gray

    if ($Building) {
        Write-Host "Building filter: $Building" -ForegroundColor Gray
    }

    if ($Capacity -gt 0) {
        Write-Warning "Capacity filtering is not supported by Outlook room list data in this script. -Capacity $Capacity will be ignored."
    }

    Write-Host ""

    # Try to get room lists from Exchange
    # Note: GetRoomLists() requires Exchange 2010+ or Office 365
    $rooms = @()
    $roomListFound = $false

    try {
        # First, try to get room lists
        $addressLists = $Namespace.AddressLists
        $roomList = $null

        foreach ($list in $addressLists) {
            if ($list.Name -like "*Room*" -or $list.Name -like "*Conference*") {
                $roomList = $list
                break
            }
        }

        if ($roomList) {
            $roomListFound = $true
            Write-Host "Found room list: $($roomList.Name)" -ForegroundColor Cyan
            Write-Host ""

            foreach ($entry in $roomList.AddressEntries) {
                $room = @{
                    Name = $entry.Name
                    Address = $entry.Address
                    Available = $true
                }

                # Filter by building name if specified
                if ($Building -and -not ($room.Name -like "*$Building*")) {
                    continue
                }

                # Check availability using FreeBusy
                try {
                    $recipient = $Namespace.CreateRecipient($entry.Address)
                    $recipient.Resolve()

                    if ($recipient.Resolved) {
                        $duration = [int](($End - $Start).TotalMinutes)
                        $freeBusy = $recipient.FreeBusy($Start, $freeBusyIntervalMinutes, $true)

                        # Check if any slot is busy
                        $numSlots = [math]::Ceiling($duration / $freeBusyIntervalMinutes)
                        for ($i = 0; $i -lt $numSlots; $i++) {
                            if ($i -lt $freeBusy.Length) {
                                if ($freeBusy[$i] -ne '0') {
                                    $room.Available = $false
                                    break
                                }
                            }
                        }
                    }
                } catch { }

                $rooms += $room
            }
        } else {
            Write-Host "No room list found in your organization." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "To use this feature, your Exchange administrator must:" -ForegroundColor Gray
            Write-Host "  1. Create room mailboxes for conference rooms" -ForegroundColor Gray
            Write-Host "  2. Create a room list to group them" -ForegroundColor Gray
            Write-Host ""
            Write-Host "Alternative: Manually specify room email addresses in meeting invites." -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "Unable to retrieve room list from Exchange." -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "This feature requires:" -ForegroundColor Gray
        Write-Host "  - Microsoft Exchange Server or Office 365" -ForegroundColor Gray
        Write-Host "  - Room mailboxes configured by your administrator" -ForegroundColor Gray
        Write-Host "  - Proper permissions to view room calendars" -ForegroundColor Gray
    }

    # Display results (only if we found a room list)
    if ($roomListFound -and $rooms.Count -gt 0) {
        $availableRooms = $rooms | Where-Object { $_.Available }
        $unavailableRooms = $rooms | Where-Object { -not $_.Available }

        if ($availableRooms.Count -gt 0) {
            Write-Host "--- Available Rooms ($($availableRooms.Count)) ---" -ForegroundColor Green
            $count = 0
            foreach ($room in $availableRooms) {
                $count++
                Write-Host "$count. $($room.Name)" -ForegroundColor Green
                Write-Host "      $($room.Address)" -ForegroundColor Gray
            }
            Write-Host ""
        }

        if ($unavailableRooms.Count -gt 0) {
            Write-Host "--- Unavailable Rooms ($($unavailableRooms.Count)) ---" -ForegroundColor Red
            $count = 0
            foreach ($room in $unavailableRooms) {
                $count++
                Write-Host "$count. $($room.Name)" -ForegroundColor Red
                Write-Host "      $($room.Address)" -ForegroundColor Gray
            }
            Write-Host ""
        }

        Write-Host "To book a room, include its address in meeting attendees:" -ForegroundColor Cyan
        Write-Host "outlook-calendar-create.ps1 -Subject 'Meeting' -Start '$($Start.ToString('yyyy-MM-dd HH:mm'))' -Attendees 'room@example.com'" -ForegroundColor Gray
    } elseif ($roomListFound) {
        Write-Host "No rooms found matching criteria." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Namespace) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
