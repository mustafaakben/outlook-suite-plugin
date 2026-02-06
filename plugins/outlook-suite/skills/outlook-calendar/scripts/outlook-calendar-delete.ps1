[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [int]$Days = 7
)

# Outlook Calendar Delete Script
# Usage: .\outlook-calendar-delete.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-calendar-delete.ps1 -Index 1 (fallback)
# Specify date range: .\outlook-calendar-delete.ps1 -Index 3 -Days 14

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
        Write-Host "Usage: .\outlook-calendar-delete.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if ($targetEvent) {
        # Store event info before deleting
        $subject = $targetEvent.Subject
        $startTime = $targetEvent.Start
        $location = $targetEvent.Location

        # Check if it's a meeting with attendees
        $isMeeting = $targetEvent.MeetingStatus -ne 0

        $targetDescription = "'$subject' on $($startTime.ToString('yyyy-MM-dd HH:mm'))"
        if ($PSCmdlet.ShouldProcess($targetDescription, "Delete calendar event")) {
            # Delete the event
            $targetEvent.Delete()

            Write-Host "`n=== EVENT DELETED ===" -ForegroundColor Green
            Write-Host "Subject: $subject" -ForegroundColor Yellow
            Write-Host "Date: $($startTime.ToString('ddd, MMM d, yyyy'))"
            Write-Host "Time: $($startTime.ToString('h:mm tt'))" -ForegroundColor Gray

            if ($location) {
                Write-Host "Location: $location" -ForegroundColor Gray
            }

            if ($isMeeting) {
                Write-Host "`nNote: This was a meeting. Cancellation may be sent to attendees." -ForegroundColor Yellow
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
