param(
    [int]$Limit = 20,
    [switch]$IncludeCompleted,
    [ValidateSet("All", "NotStarted", "InProgress", "Completed", "Waiting", "Deferred")]
    [string]$Status = "All"
)

# Outlook Tasks List Script
# Usage: .\outlook-tasks-list.ps1
# Include completed: .\outlook-tasks-list.ps1 -IncludeCompleted
# Filter by status: .\outlook-tasks-list.ps1 -Status InProgress

# Status constants:
# 0 = olTaskNotStarted
# 1 = olTaskInProgress
# 2 = olTaskComplete
# 3 = olTaskWaiting
# 4 = olTaskDeferred

$statusMap = @{
    "NotStarted" = 0
    "InProgress" = 1
    "Completed" = 2
    "Waiting" = 3
    "Deferred" = 4
}

$statusNames = @{
    0 = "Not Started"
    1 = "In Progress"
    2 = "Completed"
    3 = "Waiting"
    4 = "Deferred"
}

$priorityNames = @{
    0 = "Low"
    1 = "Normal"
    2 = "High"
}

$Outlook = $null
$Namespace = $null
$Tasks = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")
    $Tasks = $Namespace.GetDefaultFolder(13)  # olFolderTasks = 13

    $items = $Tasks.Items
    $items.Sort("[DueDate]")

    Write-Host "`n=== OUTLOOK TASKS ===" -ForegroundColor Cyan

    if ($Status -ne "All") {
        Write-Host "(Status: $Status)" -ForegroundColor Gray
    } elseif (-not $IncludeCompleted) {
        Write-Host "(Excluding completed tasks)" -ForegroundColor Gray
    }

    Write-Host ""

    $count = 0
    $displayed = 0

    foreach ($task in $items) {
        $count++

        # Filter by status
        if ($Status -ne "All") {
            if ($task.Status -ne $statusMap[$Status]) {
                continue
            }
        } elseif (-not $IncludeCompleted) {
            if ($task.Status -eq 2) {  # Skip completed
                continue
            }
        }

        $displayed++
        if ($displayed -gt $Limit) { break }

        # Display task
        $subject = if ($task.Subject) { $task.Subject } else { "(No Subject)" }
        $taskStatus = $statusNames[$task.Status]
        $priority = $priorityNames[$task.Importance]

        # Color based on status
        $color = switch ($task.Status) {
            0 { "Yellow" }      # Not Started
            1 { "Cyan" }        # In Progress
            2 { "Green" }       # Completed
            3 { "Gray" }        # Waiting
            4 { "DarkGray" }    # Deferred
            default { "White" }
        }

        Write-Host "$displayed. $subject" -ForegroundColor $color
        Write-Host "      EntryID: $($task.EntryID)" -ForegroundColor DarkGray

        # Status and priority
        $statusLine = "      Status: $taskStatus"
        if ($priority -ne "Normal") {
            $statusLine += " | Priority: $priority"
        }
        Write-Host $statusLine -ForegroundColor Gray

        # Due date
        if ($task.DueDate -and $task.DueDate -ne "1/1/4501") {
            $dueDate = $task.DueDate
            $today = Get-Date
            $daysUntil = ($dueDate - $today).Days

            $dueColor = "Gray"
            $dueLabel = ""
            if ($daysUntil -lt 0) {
                $dueColor = "Red"
                $dueLabel = " (OVERDUE)"
            } elseif ($daysUntil -eq 0) {
                $dueColor = "Yellow"
                $dueLabel = " (TODAY)"
            } elseif ($daysUntil -eq 1) {
                $dueLabel = " (Tomorrow)"
            }

            Write-Host "      Due: $($dueDate.ToString('ddd, MMM d, yyyy'))$dueLabel" -ForegroundColor $dueColor
        }

        # Percent complete if in progress
        if ($task.Status -eq 1 -and $task.PercentComplete -gt 0) {
            Write-Host "      Progress: $($task.PercentComplete)%" -ForegroundColor Gray
        }

        Write-Host ""
    }

    if ($displayed -eq 0) {
        Write-Host "No tasks found." -ForegroundColor Gray
    } else {
        Write-Host "--- Showing $displayed tasks ---" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Tasks) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Tasks) | Out-Null } catch {}
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
