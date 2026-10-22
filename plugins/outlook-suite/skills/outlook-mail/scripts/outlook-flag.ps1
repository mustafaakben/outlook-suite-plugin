param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [ValidateSet("NoDate", "Today", "Tomorrow", "ThisWeek", "NextWeek")]
    [string]$When = "NoDate",
    [string]$Flag = "Follow up",
    [datetime]$StartDate,
    [datetime]$DueDate,
    [datetime]$ReminderDate,
    [switch]$Json
)

# Outlook Flag: flag an email for follow-up (Outlook's task flag: MarkAsTask).
# Usage: .\outlook-flag.ps1 -EntryID "00000000..."
# Usage: .\outlook-flag.ps1 -EntryID "00000000..." -When Tomorrow
# Usage: .\outlook-flag.ps1 -EntryID "00000000..." -DueDate "2026-11-01" -ReminderDate "2026-10-31 09:00"

. (Join-Path $PSScriptRoot "_common.ps1")

$intervals = @{ "Today" = 0; "Tomorrow" = 1; "ThisWeek" = 2; "NextWeek" = 3; "NoDate" = 4 }

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    try {
        $email.MarkAsTask($intervals[$When])
    } catch {
        # Outlook refuses MarkAsTask on drafts; set the flag properties directly instead.
        $email.FlagStatus = 2
        if ($When -ne "NoDate" -and -not $DueDate) {
            $days = @{ "Today" = 0; "Tomorrow" = 1; "ThisWeek" = [int](5 - [int](Get-Date).DayOfWeek); "NextWeek" = [int](12 - [int](Get-Date).DayOfWeek) }[$When]
            $DueDate = (Get-Date).Date.AddDays([Math]::Max(0, $days))
        }
    }
    if ($Flag) { $email.FlagRequest = $Flag }
    if ($StartDate) { $email.TaskStartDate = $StartDate }
    if ($DueDate) {
        if (-not $StartDate -and ([datetime]$email.TaskStartDate).Year -ge 4500) { $email.TaskStartDate = (Get-Date).Date }
        $email.TaskDueDate = $DueDate
    }
    if ($ReminderDate) { $email.ReminderSet = $true; $email.ReminderTime = $ReminderDate }
    $email.Save()

    $changes = [ordered]@{ Flag = $Flag; When = $When }
    if (([datetime]$email.TaskStartDate).Year -lt 4500) { $changes.Start = Format-MailDate $email.TaskStartDate }
    if (([datetime]$email.TaskDueDate).Year -lt 4500) { $changes.Due = Format-MailDate $email.TaskDueDate }
    if ($ReminderDate) { $changes.Reminder = Format-MailDate $ReminderDate }
    Write-ActionResult -Title "EMAIL FLAGGED" -Item $email -Changes $changes -Json:$Json -Action "flag"
}
catch {
    Stop-WithError $_
}
