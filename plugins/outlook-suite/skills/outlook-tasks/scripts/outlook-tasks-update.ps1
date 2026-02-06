param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$Subject = "",
    [datetime]$DueDate,
    [datetime]$StartDate,
    [AllowNull()]
    [ValidateSet("NotStarted", "InProgress", "Completed", "Waiting", "Deferred")]
    [string]$Status = $null,
    [AllowNull()]
    [ValidateSet("Low", "Normal", "High")]
    [string]$Priority = $null,
    [string]$Body = "",
    [int]$PercentComplete = -1,
    [datetime]$Reminder,
    [switch]$ClearReminder,
    [switch]$IncludeCompleted
)

# Outlook Task Update Script
# Usage: .\outlook-tasks-update.ps1 -EntryID "00000000..." -Subject "New Title" (preferred)
# Usage: .\outlook-tasks-update.ps1 -Index 1 -Subject "New Title" (fallback)
# Update status: .\outlook-tasks-update.ps1 -EntryID "00000000..." -Status InProgress -PercentComplete 50
# Update due date: .\outlook-tasks-update.ps1 -EntryID "00000000..." -DueDate "2026-02-15"

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

$priorityMap = @{
    "Low" = 0
    "Normal" = 1
    "High" = 2
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
    $statusProvided = $PSBoundParameters.ContainsKey("Status") -and $null -ne $Status
    $priorityProvided = $PSBoundParameters.ContainsKey("Priority") -and $null -ne $Priority
    $percentProvided = $PSBoundParameters.ContainsKey("PercentComplete")
    $reminderProvided = $PSBoundParameters.ContainsKey("Reminder")

    if ($reminderProvided -and $ClearReminder) {
        Write-Host "ERROR: -Reminder and -ClearReminder cannot be used together." -ForegroundColor Red
        Write-Host "Use -Reminder to set a reminder, or -ClearReminder to remove it." -ForegroundColor Gray
        return
    }

    if ($percentProvided -and ($PercentComplete -lt 0 -or $PercentComplete -gt 100)) {
        Write-Host "ERROR: -PercentComplete must be between 0 and 100." -ForegroundColor Red
        return
    }

    if ($statusProvided -and $percentProvided) {
        $statusPercentConflict = $false
        $statusPercentGuidance = ""

        switch ($Status) {
            "Completed" {
                if ($PercentComplete -ne 100) {
                    $statusPercentConflict = $true
                    $statusPercentGuidance = "Use -Status Completed with -PercentComplete 100."
                }
            }
            "NotStarted" {
                if ($PercentComplete -ne 0) {
                    $statusPercentConflict = $true
                    $statusPercentGuidance = "Use -Status NotStarted with -PercentComplete 0."
                }
            }
            "InProgress" {
                if ($PercentComplete -le 0 -or $PercentComplete -ge 100) {
                    $statusPercentConflict = $true
                    $statusPercentGuidance = "Use -Status InProgress with -PercentComplete between 1 and 99."
                }
            }
        }

        if ($statusPercentConflict) {
            Write-Host "ERROR: -Status '$Status' conflicts with -PercentComplete $PercentComplete." -ForegroundColor Red
            Write-Host $statusPercentGuidance -ForegroundColor Gray
            return
        }
    }

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
        Write-Host "Usage: .\outlook-tasks-update.ps1 -EntryID ""00000000..."" -Subject ""New Title""" -ForegroundColor Gray
    }

    if ($targetTask) {
        # Track changes
        $changes = @()
        $oldSubject = $targetTask.Subject

        # Update fields if provided
        if ($Subject) {
            $targetTask.Subject = $Subject
            $changes += "Subject: '$oldSubject' -> '$Subject'"
        }

        if ($DueDate) {
            $oldDue = if ($targetTask.DueDate -and $targetTask.DueDate -ne "1/1/4501") {
                $targetTask.DueDate.ToString("g")
            } else {
                "(none)"
            }
            $targetTask.DueDate = $DueDate
            $changes += "Due Date: $oldDue -> $($DueDate.ToString('g'))"
        }

        if ($StartDate) {
            $oldStart = if ($targetTask.StartDate -and $targetTask.StartDate -ne "1/1/4501") {
                $targetTask.StartDate.ToString("g")
            } else {
                "(none)"
            }
            $targetTask.StartDate = $StartDate
            $changes += "Start Date: $oldStart -> $($StartDate.ToString('g'))"
        }

        if ($statusProvided) {
            $oldStatus = $statusNames[$targetTask.Status]
            $targetTask.Status = $statusMap[$Status]
            $changes += "Status: '$oldStatus' -> '$Status'"

            # If marking complete, set percent to 100 unless explicitly provided
            if ($Status -eq "Completed" -and -not $percentProvided) {
                $oldPercent = $targetTask.PercentComplete
                $targetTask.PercentComplete = 100
                $changes += "Progress: $oldPercent% -> 100%"
            }
        }

        if ($priorityProvided) {
            $oldPriority = $priorityNames[$targetTask.Importance]
            $targetTask.Importance = $priorityMap[$Priority]
            $changes += "Priority: '$oldPriority' -> '$Priority'"
        }

        if ($percentProvided) {
            $oldPercent = $targetTask.PercentComplete
            $targetTask.PercentComplete = $PercentComplete
            $changes += "Progress: $oldPercent% -> $PercentComplete%"

            if (-not $statusProvided) {
                # Auto-update status based on percent only when status wasn't explicitly provided
                if ($PercentComplete -eq 100 -and $targetTask.Status -ne 2) {
                    $targetTask.Status = 2
                    $changes += "Status: Auto-set to 'Completed'"
                } elseif ($PercentComplete -gt 0 -and $PercentComplete -lt 100 -and $targetTask.Status -eq 0) {
                    $targetTask.Status = 1
                    $changes += "Status: Auto-set to 'In Progress'"
                }
            }
        }

        if ($Body) {
            $targetTask.Body = $Body
            $changes += "Body updated"
        }

        if ($reminderProvided) {
            $targetTask.ReminderSet = $true
            $targetTask.ReminderTime = $Reminder
            $changes += "Reminder: Set for $($Reminder.ToString('g'))"
        }

        if ($ClearReminder) {
            $targetTask.ReminderSet = $false
            $changes += "Reminder: Cleared"
        }

        if ($changes.Count -eq 0) {
            Write-Host "`nNo changes specified. Use parameters like -Subject, -DueDate, -Status, etc." -ForegroundColor Yellow
        } else {
            # Save changes
            $targetTask.Save()

            Write-Host "`n=== TASK UPDATED ===" -ForegroundColor Green
            Write-Host "Task: $($targetTask.Subject)" -ForegroundColor Yellow
            Write-Host "Status: $($statusNames[$targetTask.Status])" -ForegroundColor Gray

            if ($targetTask.DueDate -and $targetTask.DueDate -ne "1/1/4501") {
                Write-Host "Due: $($targetTask.DueDate.ToString('ddd, MMM d, yyyy'))" -ForegroundColor Gray
            }

            if ($targetTask.PercentComplete -gt 0 -and $targetTask.PercentComplete -lt 100) {
                Write-Host "Progress: $($targetTask.PercentComplete)%" -ForegroundColor Gray
            }

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
