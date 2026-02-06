param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [int]$Days = 7
)

# Outlook Calendar Attendees Script
# View meeting responses and attendee status
# Usage: .\outlook-calendar-attendees.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-calendar-attendees.ps1 -Index 1 (fallback)
# With date range: .\outlook-calendar-attendees.ps1 -Index 1 -Days 14

# Meeting response constants
# olResponseNone = 0
# olResponseOrganizer = 1
# olResponseTentative = 2
# olResponseAccepted = 3
# olResponseDeclined = 4
# olResponseNotResponded = 5

$responseNames = @{
    0 = "None"
    1 = "Organizer"
    2 = "Tentative"
    3 = "Accepted"
    4 = "Declined"
    5 = "Not Responded"
}

$responseColors = @{
    0 = "Gray"
    1 = "Cyan"
    2 = "Yellow"
    3 = "Green"
    4 = "Red"
    5 = "DarkGray"
}

# Attendee type constants
# olRequired = 1
# olOptional = 2
# olResource = 3

$attendeeTypes = @{
    1 = "Required"
    2 = "Optional"
    3 = "Resource"
}

$Outlook = $null
$Namespace = $null
$Calendar = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")
    $targetEvent = $null

    if ($EntryID) {
        $targetEvent = $Namespace.GetItemFromID($EntryID)
        if (-not $targetEvent) {
            Write-Host "ERROR: Event not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-calendar-list.ps1 to find events and get EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
        $Calendar = $Namespace.GetDefaultFolder(9)  # olFolderCalendar = 9

        $items = $Calendar.Items
        $items.Sort("[Start]")
        $items.IncludeRecurrences = $true

        # Get events from now to Days ahead (fixed Outlook Restrict format)
        $now = Get-Date
        $restrictCulture = [System.Globalization.CultureInfo]::InvariantCulture
        $startDate = $now.ToString("MM/dd/yyyy HH:mm", $restrictCulture)
        $endDate = $now.AddDays($Days).ToString("MM/dd/yyyy HH:mm", $restrictCulture)

        $filter = "[Start] >= '$startDate' AND [Start] <= '$endDate'"
        $events = $items.Restrict($filter)

        $count = 0
        foreach ($event in $events) {
            $count++
            if ($count -eq $Index) {
                $targetEvent = $event
                break
            }
        }

        if (-not $targetEvent) {
            Write-Host "`nEvent at index $Index not found in next $Days days." -ForegroundColor Red
            Write-Host "Use outlook-calendar-list.ps1 -Days $Days to see available events." -ForegroundColor Gray
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target an event." -ForegroundColor Red
        Write-Host "Usage: .\outlook-calendar-attendees.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if ($targetEvent) {
        Write-Host "`n=== MEETING ATTENDEES ===" -ForegroundColor Cyan
        Write-Host "Subject: $($targetEvent.Subject)" -ForegroundColor Yellow
        Write-Host "Date: $($targetEvent.Start.ToString('ddd, MMM d, yyyy'))"
        Write-Host "Time: $($targetEvent.Start.ToString('h:mm tt')) - $($targetEvent.End.ToString('h:mm tt'))" -ForegroundColor Gray

        if ($targetEvent.Location) {
            Write-Host "Location: $($targetEvent.Location)" -ForegroundColor Gray
        }

        # Check if this is a meeting (has attendees)
        $recipients = $targetEvent.Recipients

        if ($recipients.Count -eq 0) {
            Write-Host "`nThis is an appointment (no attendees)." -ForegroundColor Gray
        } else {
            # Organizer
            Write-Host "`nOrganizer: $($targetEvent.Organizer)" -ForegroundColor Cyan

            # Count responses
            $accepted = 0
            $tentative = 0
            $declined = 0
            $noResponse = 0

            Write-Host "`n--- Attendees ($($recipients.Count)) ---" -ForegroundColor Cyan

            # Separate required and optional attendees
            $requiredAttendees = @()
            $optionalAttendees = @()
            $resources = @()

            foreach ($recipient in $recipients) {
                $attendee = @{
                    Name = $recipient.Name
                    Address = $recipient.Address
                    Response = $recipient.MeetingResponseStatus
                    Type = $recipient.Type
                }

                switch ($recipient.Type) {
                    1 { $requiredAttendees += $attendee }
                    2 { $optionalAttendees += $attendee }
                    3 { $resources += $attendee }
                }

                # Count responses
                switch ($recipient.MeetingResponseStatus) {
                    3 { $accepted++ }
                    2 { $tentative++ }
                    4 { $declined++ }
                    { $_ -eq 0 -or $_ -eq 5 } { $noResponse++ }
                }
            }

            # Display required attendees
            if ($requiredAttendees.Count -gt 0) {
                Write-Host "`nRequired:" -ForegroundColor White
                foreach ($att in $requiredAttendees) {
                    $status = $responseNames[$att.Response]
                    $color = $responseColors[$att.Response]
                    Write-Host "  $($att.Name)" -NoNewline
                    Write-Host " - $status" -ForegroundColor $color
                }
            }

            # Display optional attendees
            if ($optionalAttendees.Count -gt 0) {
                Write-Host "`nOptional:" -ForegroundColor Gray
                foreach ($att in $optionalAttendees) {
                    $status = $responseNames[$att.Response]
                    $color = $responseColors[$att.Response]
                    Write-Host "  $($att.Name)" -NoNewline
                    Write-Host " - $status" -ForegroundColor $color
                }
            }

            # Display resources
            if ($resources.Count -gt 0) {
                Write-Host "`nResources:" -ForegroundColor Gray
                foreach ($att in $resources) {
                    $status = $responseNames[$att.Response]
                    $color = $responseColors[$att.Response]
                    Write-Host "  $($att.Name)" -NoNewline
                    Write-Host " - $status" -ForegroundColor $color
                }
            }

            # Summary
            Write-Host "`n--- Response Summary ---" -ForegroundColor Cyan
            Write-Host "Accepted: $accepted" -ForegroundColor Green
            Write-Host "Tentative: $tentative" -ForegroundColor Yellow
            Write-Host "Declined: $declined" -ForegroundColor Red
            Write-Host "No Response: $noResponse" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Calendar) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Calendar) | Out-Null } catch {}
    }
    if ($Namespace) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
