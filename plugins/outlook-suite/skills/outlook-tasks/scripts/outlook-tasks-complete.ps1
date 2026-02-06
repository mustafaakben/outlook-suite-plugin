param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [string]$Subject = ""
)

# Outlook Task Complete Script
# Usage: .\outlook-tasks-complete.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-tasks-complete.ps1 -Index 1 (fallback)
# By subject: .\outlook-tasks-complete.ps1 -Subject "Finish report"

$Outlook = $null
$Namespace = $null
$Tasks = $null

try {
    if (-not $EntryID -and $Index -eq 0 -and -not $Subject) {
        Write-Host "`nPlease specify -EntryID, -Index, or -Subject to complete a task:" -ForegroundColor Red
        Write-Host "  -EntryID ""00000000...""" -ForegroundColor Gray
        Write-Host "  -Index 1" -ForegroundColor Gray
        Write-Host "  -Subject ""Finish report""" -ForegroundColor Gray
    } else {
        # Connect to Outlook - try active instance first
        try {
            $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            $Outlook = New-Object -ComObject Outlook.Application
            Start-Sleep -Milliseconds 500
        }

        $Namespace = $Outlook.GetNamespace("MAPI")
        $targetTask = $null
        $taskSubject = ""

        if ($EntryID) {
            $targetTask = $Namespace.GetItemFromID($EntryID)
            if (-not $targetTask) {
                Write-Host "ERROR: Task not found with provided EntryID." -ForegroundColor Red
                Write-Host "Tip: Use outlook-tasks-list.ps1 to find tasks and get EntryIDs." -ForegroundColor Gray
            } else {
                $taskSubject = $targetTask.Subject
            }
        } elseif ($Index -gt 0) {
            $Tasks = $Namespace.GetDefaultFolder(13)  # olFolderTasks = 13
            $items = $Tasks.Items
            $items.Sort("[DueDate]")

            # Find by index (only counting non-completed tasks)
            $count = 0
            foreach ($task in $items) {
                if ($task.Status -ne 2) {  # Skip completed
                    $count++
                    if ($count -eq $Index) {
                        $targetTask = $task
                        $taskSubject = $task.Subject
                        break
                    }
                }
            }

            if (-not $targetTask) {
                Write-Host "`nTask at index $Index not found." -ForegroundColor Red
                Write-Host "Use outlook-tasks-list.ps1 to see available tasks." -ForegroundColor Gray
            }
        } elseif ($Subject) {
            $Tasks = $Namespace.GetDefaultFolder(13)  # olFolderTasks = 13
            $items = $Tasks.Items
            $items.Sort("[DueDate]")

            $exactMatchCount = 0
            foreach ($task in $items) {
                if ($task.Subject -and $task.Subject -eq $Subject) {
                    $exactMatchCount++
                    if ($exactMatchCount -eq 1) {
                        $targetTask = $task
                        $taskSubject = $task.Subject
                    }
                }
            }

            if ($exactMatchCount -eq 0) {
                Write-Host "`nTask '$Subject' not found." -ForegroundColor Red
            } elseif ($exactMatchCount -gt 1) {
                $targetTask = $null
                Write-Host "`nMultiple tasks found with exact subject '$Subject' ($exactMatchCount matches)." -ForegroundColor Red
                Write-Host "Use -EntryID or -Index to target a single task." -ForegroundColor Gray
                Write-Host "Tip: Run outlook-tasks-list.ps1 to find the correct task." -ForegroundColor Gray
            }
        }

        if ($targetTask) {
            # Check if already completed
            if ($targetTask.Status -eq 2) {
                Write-Host "`nTask is already completed." -ForegroundColor Yellow
                Write-Host "Subject: $taskSubject" -ForegroundColor Gray
            } else {
                # Mark as complete
                $targetTask.MarkComplete()

                Write-Host "`n=== TASK COMPLETED ===" -ForegroundColor Green
                Write-Host "Subject: $taskSubject" -ForegroundColor Yellow
                Write-Host "Status: Completed" -ForegroundColor Gray
                Write-Host "Completed: $((Get-Date).ToString('g'))" -ForegroundColor Gray
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
