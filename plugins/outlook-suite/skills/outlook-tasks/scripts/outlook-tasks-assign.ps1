param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [string]$AssignTo,

    [switch]$KeepCopy,
    [switch]$IncludeCompleted
)

# Outlook Task Assign Script
# Delegate/assign a task to another person
# Usage: .\outlook-tasks-assign.ps1 -EntryID "00000000..." -AssignTo "colleague@example.com" (preferred)
# Usage: .\outlook-tasks-assign.ps1 -Index 1 -AssignTo "colleague@example.com" (fallback)
# Keep copy: .\outlook-tasks-assign.ps1 -EntryID "00000000..." -AssignTo "colleague@example.com" -KeepCopy

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
    $targetTask = $null

    if ($EntryID) {
        $targetTask = $Namespace.GetItemFromID($EntryID)
        if (-not $targetTask) {
            Write-Host "ERROR: Task not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-tasks-list.ps1 to find tasks and get EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
        $Tasks = $Namespace.GetDefaultFolder(13)  # olFolderTasks = 13
        $items = $Tasks.Items
        $items.Sort("[DueDate]")

        # Build filtered list
        $filteredTasks = @()
        foreach ($task in $items) {
            if (-not $IncludeCompleted -and $task.Status -eq 2) {
                continue
            }
            $filteredTasks += $task
        }

        # Find the task at the specified index
        if ($Index -lt 1 -or $Index -gt $filteredTasks.Count) {
            Write-Host "`nTask at index $Index not found." -ForegroundColor Red
            Write-Host "Use outlook-tasks-list.ps1 to see available tasks." -ForegroundColor Gray
        } else {
            $targetTask = $filteredTasks[$Index - 1]
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target a task." -ForegroundColor Red
        Write-Host "Usage: .\outlook-tasks-assign.ps1 -EntryID ""00000000..."" -AssignTo ""colleague@example.com""" -ForegroundColor Gray
    }

    if ($targetTask) {
        $taskSubject = $targetTask.Subject

        # Check if task is already assigned
        if ($targetTask.Delegator) {
            Write-Host "`nWarning: This task was already assigned by: $($targetTask.Delegator)" -ForegroundColor Yellow
        }

        # Assign the task
        $taskRequest = $targetTask.Assign()

        # Add the recipient
        $recipient = $taskRequest.Recipients.Add($AssignTo)
        $recipient.Resolve()

        if (-not $recipient.Resolved) {
            Write-Host "`nCould not resolve recipient: $AssignTo" -ForegroundColor Red
            Write-Host "Please verify the email address." -ForegroundColor Gray
        } else {
            # Set whether to keep a copy
            if ($KeepCopy) {
                $taskRequest.KeepAfterAssignment = $true
            }

            # Send the task request
            $taskRequest.Send()

            Write-Host "`n=== TASK ASSIGNED ===" -ForegroundColor Green
            Write-Host "Task: $taskSubject" -ForegroundColor Yellow
            Write-Host "Assigned to: $AssignTo" -ForegroundColor Cyan

            if ($targetTask.DueDate -and $targetTask.DueDate -ne "1/1/4501") {
                Write-Host "Due: $($targetTask.DueDate.ToString('ddd, MMM d, yyyy'))" -ForegroundColor Gray
            }

            if ($KeepCopy) {
                Write-Host "`nA copy will be kept in your task list." -ForegroundColor Gray
            } else {
                Write-Host "`nTask will be removed from your list once accepted." -ForegroundColor Gray
            }

            Write-Host "`nTask request sent. The recipient will receive an email to accept/decline." -ForegroundColor Cyan
        }
    }
}
catch {
    if ($_.Exception.Message -like "*Operation aborted*") {
        Write-Host "`nTask assignment was cancelled." -ForegroundColor Yellow
    } else {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "`nNote: Task assignment requires Exchange or Microsoft 365." -ForegroundColor Gray
    }
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
