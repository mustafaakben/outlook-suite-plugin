param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Accept", "Tentative", "Decline")]
    [string]$Response,

    [int]$Days = 7,
    [switch]$NoResponse
)

# Outlook Calendar Respond Script
# Usage: .\outlook-calendar-respond.ps1 -EntryID "00000000..." -Response Accept (preferred)
# Usage: .\outlook-calendar-respond.ps1 -Index 1 -Response Accept (fallback)
# Decline: .\outlook-calendar-respond.ps1 -EntryID "00000000..." -Response Decline
# Tentative without sending response: .\outlook-calendar-respond.ps1 -Index 1 -Response Tentative -NoResponse

# Response constants
# olMeetingAccepted = 3
# olMeetingTentative = 2
# olMeetingDeclined = 4

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
        Write-Host "Usage: .\outlook-calendar-respond.ps1 -EntryID ""00000000..."" -Response Accept" -ForegroundColor Gray
    }

    if ($targetEvent) {
        # Check if it's actually a meeting (not just an appointment)
        if ($targetEvent.MeetingStatus -eq 0) {
            Write-Host "`nThis is an appointment, not a meeting invite." -ForegroundColor Yellow
            Write-Host "You can only respond to meetings you've been invited to." -ForegroundColor Gray
            Write-Host "Subject: $($targetEvent.Subject)" -ForegroundColor Gray
        } else {
            # Store event info
            $subject = $targetEvent.Subject
            $organizer = $targetEvent.Organizer
            $startTime = $targetEvent.Start

            # Map response to constant
            $responseType = switch ($Response) {
                "Accept" { 3 }      # olMeetingAccepted
                "Tentative" { 2 }   # olMeetingTentative
                "Decline" { 4 }     # olMeetingDeclined
            }

            # Send response (or not)
            $sendResponse = -not $NoResponse

            # Respond to the meeting
            $responseItem = $targetEvent.Respond($responseType, $sendResponse)

            # Determine color based on response
            $color = switch ($Response) {
                "Accept" { "Green" }
                "Tentative" { "Yellow" }
                "Decline" { "Red" }
            }

            Write-Host "`n=== MEETING RESPONSE: $($Response.ToUpper()) ===" -ForegroundColor $color
            Write-Host "Subject: $subject" -ForegroundColor Yellow
            Write-Host "Organizer: $organizer"
            Write-Host "Date: $($startTime.ToString('ddd, MMM d, yyyy'))"
            Write-Host "Time: $($startTime.ToString('h:mm tt'))" -ForegroundColor Gray

            if ($sendResponse) {
                Write-Host "`nResponse sent to organizer." -ForegroundColor Cyan
            } else {
                Write-Host "`nResponse saved (not sent to organizer)." -ForegroundColor Gray
            }
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
