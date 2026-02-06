param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [string]$To,

    [string]$Body = "",
    [int]$Days = 7,
    [switch]$Send
)

# Outlook Calendar Forward Script
# Forward a calendar event as a vCal attachment via ForwardAsVcal
# Usage: .\outlook-calendar-forward.ps1 -EntryID "00000000..." -To "colleague@example.com" (preferred)
# Usage: .\outlook-calendar-forward.ps1 -Index 1 -To "colleague@example.com" (fallback)
# With message: .\outlook-calendar-forward.ps1 -EntryID "00000000..." -To "email@example.com" -Body "FYI - meeting details"
# Send immediately: .\outlook-calendar-forward.ps1 -EntryID "00000000..." -To "email@example.com" -Send

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
        Write-Host "Usage: .\outlook-calendar-forward.ps1 -EntryID ""00000000..."" -To ""colleague@example.com""" -ForegroundColor Gray
    }

    if ($targetEvent) {
        # Forward the event
        # ForwardAsVcal creates a new mail item with the event as a vCal attachment
        $forwardMail = $targetEvent.ForwardAsVcal()

        $forwardMail.To = $To

        if ($Body) {
            # Prepend custom body to the generated content
            $forwardMail.Body = $Body + "`n`n" + $forwardMail.Body
        }

        if ($Send) {
            $forwardMail.Send()

            Write-Host "`n=== EVENT FORWARDED ===" -ForegroundColor Green
            Write-Host "Subject: $($targetEvent.Subject)" -ForegroundColor Yellow
            Write-Host "Date: $($targetEvent.Start.ToString('ddd, MMM d, yyyy'))"
            Write-Host "Time: $($targetEvent.Start.ToString('h:mm tt')) - $($targetEvent.End.ToString('h:mm tt'))" -ForegroundColor Gray
            Write-Host "`nForwarded to: $To" -ForegroundColor Cyan
            Write-Host "Status: Sent" -ForegroundColor Green
        } else {
            $forwardMail.Save()

            Write-Host "`n=== FORWARD DRAFT CREATED ===" -ForegroundColor Green
            Write-Host "Subject: $($targetEvent.Subject)" -ForegroundColor Yellow
            Write-Host "Date: $($targetEvent.Start.ToString('ddd, MMM d, yyyy'))"
            Write-Host "Time: $($targetEvent.Start.ToString('h:mm tt')) - $($targetEvent.End.ToString('h:mm tt'))" -ForegroundColor Gray
            Write-Host "`nTo: $To" -ForegroundColor Cyan
            Write-Host "Status: Saved as draft" -ForegroundColor Yellow

            Write-Host "`nReview in Drafts folder, then send manually or use -Send flag." -ForegroundColor Gray
        }

        # Show attachment info
        Write-Host "`nAttachment: Calendar event (vCal/.vcs)" -ForegroundColor Gray
        Write-Host "Recipients can add this event to their calendar." -ForegroundColor Gray
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
