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
    [switch]$ClearDueDate,
    [string]$Categories = $null,
    [switch]$IncludeCompleted,
    [switch]$Json
)

# Outlook Task Update: change a task's subject, dates, status, priority, progress, body,
# reminder or categories. Only the fields you pass are changed.
# Usage: .\outlook-tasks-update.ps1 -EntryID "00000000..." -Subject "New Title" (preferred)
# Usage: .\outlook-tasks-update.ps1 -Index 1 -Subject "New Title" (position in outlook-tasks-list.ps1 order)
# Status: .\outlook-tasks-update.ps1 -EntryID "00000000..." -Status InProgress -PercentComplete 50
# Dates: .\outlook-tasks-update.ps1 -EntryID "00000000..." -DueDate "2026-02-15" | -ClearDueDate
# Categories: .\outlook-tasks-update.ps1 -EntryID "00000000..." -Categories "Teaching, Urgent" (empty string clears)

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")

try {
    $statusProvided = $PSBoundParameters.ContainsKey("Status") -and $Status
    $priorityProvided = $PSBoundParameters.ContainsKey("Priority") -and $Priority
    $percentProvided = $PSBoundParameters.ContainsKey("PercentComplete")
    $reminderProvided = $PSBoundParameters.ContainsKey("Reminder")
    $categoriesProvided = $PSBoundParameters.ContainsKey("Categories")

    if ($reminderProvided -and $ClearReminder) { Stop-WithError "-Reminder and -ClearReminder cannot be used together." }
    if ($PSBoundParameters.ContainsKey("DueDate") -and $ClearDueDate) { Stop-WithError "-DueDate and -ClearDueDate cannot be used together." }
    if ($percentProvided -and ($PercentComplete -lt 0 -or $PercentComplete -gt 100)) { Stop-WithError "-PercentComplete must be between 0 and 100." }
    if ($statusProvided -and $percentProvided) {
        $guidance = switch ($Status) {
            "Completed" { if ($PercentComplete -ne 100) { "Use -Status Completed with -PercentComplete 100." } }
            "NotStarted" { if ($PercentComplete -ne 0) { "Use -Status NotStarted with -PercentComplete 0." } }
            "InProgress" { if ($PercentComplete -le 0 -or $PercentComplete -ge 100) { "Use -Status InProgress with -PercentComplete between 1 and 99." } }
        }
        if ($guidance) { Stop-WithError "-Status '$Status' conflicts with -PercentComplete $PercentComplete." -Hint $guidance }
    }

    $null = Connect-Outlook
    $task = Resolve-TaskItem -Namespace $Namespace -EntryID $EntryID -Index $Index -IncludeCompleted:$IncludeCompleted

    $changes = [ordered]@{}
    $before = ConvertTo-TaskRecord $task
    if ($Subject) { $task.Subject = $Subject; $changes.Subject = "'$($before.Subject)' -> '$Subject'" }
    if ($StartDate) { $task.StartDate = $StartDate; $changes.Start = "$(Format-TaskDate $before.Start) -> $(Format-TaskDate $StartDate)" }
    if ($DueDate) { $task.DueDate = $DueDate; $changes.Due = "$(if ($before.Due) { Format-TaskDate $before.Due } else { '(none)' }) -> $(Format-TaskDate $DueDate)" }
    if ($ClearDueDate) { $task.DueDate = [datetime]"4501-01-01"; $changes.Due = "$(Format-TaskDate $before.Due) -> (none)" }
    if ($statusProvided) {
        $task.Status = $script:TaskStatusMap[$Status]
        $changes.Status = "$($before.Status) -> $Status"
        if ($Status -eq "Completed" -and -not $percentProvided) { $task.PercentComplete = 100 }
    }
    if ($priorityProvided) { $task.Importance = $script:TaskPriorityMap[$Priority]; $changes.Priority = "$($before.Importance) -> $Priority" }
    if ($percentProvided) {
        $task.PercentComplete = $PercentComplete
        $changes.Progress = "$($before.PercentComplete)% -> $PercentComplete%"
        if (-not $statusProvided) {
            # Keep status consistent with progress when status was not given explicitly.
            if ($PercentComplete -eq 100 -and $task.Status -ne 2) { $task.Status = 2; $changes.Status = "auto-set to Completed" }
            elseif ($PercentComplete -gt 0 -and $PercentComplete -lt 100 -and $task.Status -eq 0) { $task.Status = 1; $changes.Status = "auto-set to InProgress" }
        }
    }
    if ($Body) { $task.Body = $Body; $changes.Body = "updated" }
    if ($reminderProvided) { $task.ReminderSet = $true; $task.ReminderTime = $Reminder; $changes.Reminder = Format-MailDate $Reminder }
    if ($ClearReminder) { $task.ReminderSet = $false; $changes.Reminder = "cleared" }
    if ($categoriesProvided) { $task.Categories = [string]$Categories; $changes.Categories = "'$($before.Categories)' -> '$Categories'" }

    if ($changes.Count -eq 0) {
        Stop-WithError "No changes specified." -Hint "Pass -Subject, -DueDate, -Status, -Priority, -PercentComplete, -Body, -Reminder, -Categories, ..."
    }
    $task.Save()
    Write-TaskResult -Title "TASK UPDATED" -Task $task -Changes $changes -Json:$Json -Action "update"
}
catch {
    Stop-WithError $_
}
