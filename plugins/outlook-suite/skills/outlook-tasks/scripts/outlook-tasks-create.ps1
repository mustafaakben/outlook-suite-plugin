param(
    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [datetime]$DueDate,
    [datetime]$StartDate,
    [ValidateSet("Low", "Normal", "High")]
    [string]$Priority = "Normal",
    [string]$Body = "",
    [datetime]$Reminder
)

# Outlook Task Create Script
# Usage: .\outlook-tasks-create.ps1 -Subject "Finish report"
# With due date: .\outlook-tasks-create.ps1 -Subject "Finish report" -DueDate "2026-02-10"
# High priority: .\outlook-tasks-create.ps1 -Subject "Urgent task" -Priority High -DueDate "2026-02-05"
# With reminder: .\outlook-tasks-create.ps1 -Subject "Call client" -Reminder "2026-02-05 09:00"

$priorityMap = @{
    "Low" = 0
    "Normal" = 1
    "High" = 2
}

$Outlook = $null
$task = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    # Create task item (olTaskItem = 3)
    $task = $Outlook.CreateItem(3)

    $task.Subject = $Subject
    $task.Importance = $priorityMap[$Priority]

    if ($Body) {
        $task.Body = $Body
    }

    if ($DueDate) {
        $task.DueDate = $DueDate
    }

    if ($StartDate) {
        $task.StartDate = $StartDate
    }

    if ($Reminder) {
        $task.ReminderSet = $true
        $task.ReminderTime = $Reminder
    }

    $task.Save()

    Write-Host "`n=== TASK CREATED ===" -ForegroundColor Green
    Write-Host "Subject: $Subject" -ForegroundColor Yellow

    if ($Priority -ne "Normal") {
        Write-Host "Priority: $Priority" -ForegroundColor $(if ($Priority -eq "High") { "Red" } else { "Gray" })
    }

    if ($DueDate) {
        Write-Host "Due: $($DueDate.ToString('ddd, MMM d, yyyy'))" -ForegroundColor Gray
    }

    if ($Reminder) {
        Write-Host "Reminder: $($Reminder.ToString('g'))" -ForegroundColor Gray
    }

    Write-Host "`nTask saved to Tasks folder." -ForegroundColor Cyan
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($task) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($task) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
