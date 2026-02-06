param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [int]$Days = 7,
    [string]$Subject = "",
    [datetime]$Start,
    [datetime]$End,
    [string]$Location = "",
    [string]$Body = ""
)

# Outlook Calendar Update Script
# Usage: .\outlook-calendar-update.ps1 -EntryID "00000000..." -Subject "New Title" (preferred)
# Usage: .\outlook-calendar-update.ps1 -Index 1 -Subject "New Title" (fallback)
# Update time: .\outlook-calendar-update.ps1 -EntryID "00000000..." -Start "2026-02-05 14:00"
# Update location: .\outlook-calendar-update.ps1 -EntryID "00000000..." -Location "Room 101"

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
        Write-Host "Usage: .\outlook-calendar-update.ps1 -EntryID ""00000000..."" -Subject ""New Title""" -ForegroundColor Gray
    }

    if ($targetEvent) {
        # Track changes
        $changes = @()
        $oldSubject = $targetEvent.Subject

        # Update fields if provided
        if ($Subject) {
            $targetEvent.Subject = $Subject
            $changes += "Subject: '$oldSubject' -> '$Subject'"
        }

        if ($Start) {
            $oldStart = $targetEvent.Start
            $targetEvent.Start = $Start
            # Adjust end time to maintain duration
            $duration = ($targetEvent.End - $oldStart)
            $targetEvent.End = $Start.Add($duration)
            $changes += "Start: $($oldStart.ToString('MM/dd h:mm tt')) -> $($Start.ToString('MM/dd h:mm tt'))"
        }

        if ($End) {
            $oldEnd = $targetEvent.End
            $targetEvent.End = $End
            $changes += "End: $($oldEnd.ToString('h:mm tt')) -> $($End.ToString('h:mm tt'))"
        }

        if ($Location) {
            $oldLocation = $targetEvent.Location
            $targetEvent.Location = $Location
            $changes += "Location: '$oldLocation' -> '$Location'"
        }

        if ($Body) {
            $targetEvent.Body = $Body
            $changes += "Body updated"
        }

        if ($changes.Count -eq 0) {
            Write-Host "`nNo changes specified. Use parameters like -Subject, -Start, -Location, etc." -ForegroundColor Yellow
        } else {
            # Save changes
            $targetEvent.Save()

            Write-Host "`n=== EVENT UPDATED ===" -ForegroundColor Green
            Write-Host "Event: $($targetEvent.Subject)" -ForegroundColor Yellow
            Write-Host "Date: $($targetEvent.Start.ToString('ddd, MMM d, yyyy'))"
            Write-Host "Time: $($targetEvent.Start.ToString('h:mm tt')) - $($targetEvent.End.ToString('h:mm tt'))" -ForegroundColor Gray

            Write-Host "`nChanges made:" -ForegroundColor Cyan
            foreach ($change in $changes) {
                Write-Host "  - $change" -ForegroundColor Gray
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
