param(
    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [datetime]$DueDate,
    [datetime]$StartDate,
    [ValidateSet("Low", "Normal", "High")]
    [string]$Priority = "Normal",
    [string]$Body = "",
    [datetime]$Reminder,
    [string]$Categories = "",
    [ValidateSet("None", "Daily", "Weekly", "Monthly", "Yearly")]
    [string]$Recurrence = "None",
    [ValidateRange(1, 999)]
    [int]$Interval = 1,
    [datetime]$Until,
    [switch]$Json
)

# Outlook Task Create: save a new task in the Tasks folder (local item, nothing is sent).
# Usage: .\outlook-tasks-create.ps1 -Subject "Finish report"
# With due date: .\outlook-tasks-create.ps1 -Subject "Finish report" -DueDate "2026-02-10"
# High priority: .\outlook-tasks-create.ps1 -Subject "Urgent task" -Priority High -DueDate "2026-02-05"
# With reminder: .\outlook-tasks-create.ps1 -Subject "Call client" -Reminder "2026-02-05 09:00"
# Category: .\outlook-tasks-create.ps1 -Subject "Grade essays" -Categories "Teaching"
# Recurring: .\outlook-tasks-create.ps1 -Subject "Weekly report" -Recurrence Weekly -DueDate "2026-11-06" -Until "2026-12-18"
#   Repeats from the due date (today when no -DueDate): Weekly on that weekday, Monthly on that day
#   of the month. -Interval 2 = every 2 days/weeks/months/years. No -Until = no end date.

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")

try {
    if ($Recurrence -eq "None" -and ($PSBoundParameters.ContainsKey("Interval") -or $Until)) { Stop-WithError "-Interval and -Until need -Recurrence Daily|Weekly|Monthly|Yearly." }
    if ($Recurrence -ne "None" -and -not $DueDate) { $DueDate = (Get-Date).Date }
    if ($Until -and $DueDate -and $Until.Date -lt $DueDate.Date) { Stop-WithError "-Until is before the first due date." }

    $null = Connect-Outlook
    $task = $Outlook.CreateItem(3)   # olTaskItem
    $task.Subject = $Subject
    $task.Importance = $script:TaskPriorityMap[$Priority]
    if ($Body) { $task.Body = $Body }
    if ($StartDate) { $task.StartDate = $StartDate }
    if ($DueDate) { $task.DueDate = $DueDate }
    if ($Reminder) { $task.ReminderSet = $true; $task.ReminderTime = $Reminder }
    if ($Categories) { $task.Categories = $Categories }
    if ($Recurrence -ne "None") {
        $p = $task.GetRecurrencePattern()
        $p.RecurrenceType = @{ "Daily" = 0; "Weekly" = 1; "Monthly" = 2; "Yearly" = 5 }[$Recurrence]
        $first = $DueDate.Date
        if ($Recurrence -eq "Weekly") { $p.DayOfWeekMask = 1 -shl [int]$first.DayOfWeek }
        if ($Recurrence -eq "Monthly") { $p.DayOfMonth = $first.Day }
        if ($Recurrence -eq "Yearly") { $p.MonthOfYear = $first.Month; $p.DayOfMonth = $first.Day }
        $p.Interval = $(if ($Recurrence -eq "Yearly") { 12 * $Interval } else { $Interval })   # yearly intervals are in months
        $p.PatternStartDate = $first
        if ($Until) { $p.PatternEndDate = $Until.Date } else { $p.NoEndDate = $true }
    }
    $task.Save()

    $changes = [ordered]@{}
    if ($Reminder) { $changes.Reminder = Format-MailDate $Reminder }
    if ($Recurrence -ne "None") { $changes.Repeats = Get-TaskRecurrenceText $task }
    Write-TaskResult -Title "TASK CREATED" -Task $task -Changes $changes -Json:$Json -Action "create"
}
catch {
    Stop-WithError $_
}
