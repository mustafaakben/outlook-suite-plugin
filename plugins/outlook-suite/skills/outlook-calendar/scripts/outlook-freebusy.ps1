param(
    [Parameter(Mandatory=$true)]
    [string]$Attendees,

    [Parameter(Mandatory=$true)]
    [datetime]$Start,

    [datetime]$End,
    [ValidateRange(1, 1440)]
    [int]$IntervalMinutes = 30
)

# Outlook Free/Busy Script
# Check availability of attendees
# Usage: .\outlook-freebusy.ps1 -Attendees "john@example.com" -Start "2026-02-10 09:00"
# Multiple: .\outlook-freebusy.ps1 -Attendees "john@example.com; jane@example.com" -Start "2026-02-10 09:00" -End "2026-02-10 17:00"
# Custom interval: .\outlook-freebusy.ps1 -Attendees "john@example.com" -Start "2026-02-10 09:00" -IntervalMinutes 15

# Free/Busy constants (returned as string of numbers)
# 0 = Free
# 1 = Tentative
# 2 = Busy
# 3 = Out of Office
# 4 = Working Elsewhere (Office 365)

$statusNames = @{
    "0" = "Free"
    "1" = "Tentative"
    "2" = "Busy"
    "3" = "Out of Office"
    "4" = "Working Elsewhere"
}

$statusColors = @{
    "0" = "Green"
    "1" = "Yellow"
    "2" = "Red"
    "3" = "Magenta"
    "4" = "Cyan"
}

$statusSymbols = @{
    "0" = "[    ]"
    "1" = "[????]"
    "2" = "[BUSY]"
    "3" = "[OOF ]"
    "4" = "[ELSE]"
}

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

    # Default end time is 8 hours after start
    if (-not $End) {
        $End = $Start.AddHours(8)
    }

    if ($End -le $Start) {
        throw "End time must be later than start time."
    }

    if ($IntervalMinutes -lt 1) {
        throw "IntervalMinutes must be at least 1 minute."
    }

    # Parse attendees
    $attendeeList = $Attendees -split "[;,]" | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    Write-Host "`n=== FREE/BUSY CHECK ===" -ForegroundColor Cyan
    Write-Host "Date: $($Start.ToString('ddd, MMM d, yyyy'))"
    Write-Host "Time: $($Start.ToString('h:mm tt')) - $($End.ToString('h:mm tt'))"
    Write-Host "Interval: $IntervalMinutes minutes" -ForegroundColor Gray
    Write-Host ""

    # Calculate number of slots
    $totalMinutes = ($End - $Start).TotalMinutes
    $numSlots = [math]::Ceiling($totalMinutes / $IntervalMinutes)

    # Build time slot headers
    $timeSlots = @()
    for ($i = 0; $i -lt $numSlots; $i++) {
        $slotTime = $Start.AddMinutes($i * $IntervalMinutes)
        $timeSlots += $slotTime.ToString("h:mm")
    }

    # Display header row (time slots)
    $maxNameLength = ($attendeeList | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    if ($maxNameLength -lt 20) { $maxNameLength = 20 }

    Write-Host ("{0,-$maxNameLength}" -f "Attendee") -NoNewline -ForegroundColor Cyan
    Write-Host " | " -NoNewline

    # Show time headers (abbreviated)
    for ($i = 0; $i -lt $numSlots; $i++) {
        if ($i -lt $timeSlots.Count) {
            $slotTime = $Start.AddMinutes($i * $IntervalMinutes)
            if ($i % 2 -eq 0) {  # Show every other time to save space
                Write-Host "$($slotTime.ToString('h:mm')) " -NoNewline -ForegroundColor Gray
            } else {
                Write-Host "      " -NoNewline
            }
        }
    }
    Write-Host ""

    Write-Host ("-" * ($maxNameLength + 3 + ($numSlots * 6))) -ForegroundColor Gray

    # Get free/busy for each attendee
    foreach ($attendee in $attendeeList) {
        try {
            $recipient = $Namespace.CreateRecipient($attendee)
            $recipient.Resolve()

            if (-not $recipient.Resolved) {
                Write-Host ("{0,-$maxNameLength}" -f $attendee) -NoNewline -ForegroundColor Yellow
                Write-Host " | " -NoNewline
                Write-Host "Unable to resolve address" -ForegroundColor Red
                Write-Host ""
                continue
            }

            # Get free/busy string
            # Parameters: Start, MinPerChar (interval), CompleteFormat (include OOF/Tentative detail)
            $freeBusy = $recipient.FreeBusy($Start, $IntervalMinutes, $true)

            # Display attendee name
            $displayName = if ($attendee.Length -gt $maxNameLength) {
                $attendee.Substring(0, $maxNameLength - 3) + "..."
            } else {
                $attendee
            }
            Write-Host ("{0,-$maxNameLength}" -f $displayName) -NoNewline -ForegroundColor White
            Write-Host " | " -NoNewline

            # Display status for each slot
            for ($i = 0; $i -lt $numSlots; $i++) {
                if ($i -lt $freeBusy.Length) {
                    $status = $freeBusy[$i].ToString()
                    $symbol = switch ($status) {
                        "0" { "  .   " }  # Free
                        "1" { "  ?   " }  # Tentative
                        "2" { " BUSY " }  # Busy
                        "3" { " OOF  " }  # Out of Office
                        "4" { " ELSE " }  # Working Elsewhere
                        default { "  ?   " }
                    }
                    $color = $statusColors[$status]
                    if (-not $color) { $color = "Gray" }
                    Write-Host $symbol -NoNewline -ForegroundColor $color
                } else {
                    Write-Host "      " -NoNewline
                }
            }
            Write-Host ""
        }
        catch {
            Write-Host ("{0,-$maxNameLength}" -f $attendee) -NoNewline -ForegroundColor Yellow
            Write-Host " | " -NoNewline
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
        }
    }

    # Legend
    Write-Host ""
    Write-Host "--- Legend ---" -ForegroundColor Cyan
    Write-Host "  .   = Free" -ForegroundColor Green
    Write-Host "  ?   = Tentative" -ForegroundColor Yellow
    Write-Host " BUSY = Busy" -ForegroundColor Red
    Write-Host " OOF  = Out of Office" -ForegroundColor Magenta
    Write-Host " ELSE = Working Elsewhere" -ForegroundColor Cyan

    Write-Host "`nNote: Free/busy data requires Exchange or Microsoft 365." -ForegroundColor Gray
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
