param(
    [Parameter(Mandatory=$true)]
    [string]$EntryID,
    [string]$StoreID = "",
    [string]$Subject = "",
    [ValidateSet("NoDate", "Today", "Tomorrow", "ThisWeek", "NextWeek")]
    [string]$When = "NoDate",
    [datetime]$DueDate,
    [ValidateSet("", "Low", "Normal", "High")]
    [string]$Priority = "",
    [string]$Categories = "",
    [datetime]$Reminder,
    [switch]$AttachEmail,
    [switch]$Force,
    [switch]$Json
)

# Outlook Task From Mail: turn an email into a task that links back to it.
# The email is only read, never changed (to flag the email itself use outlook-flag.ps1).
# The task body names the sender and date and holds an outlook: link that opens the email in
# classic Outlook; the email's EntryID is also stored on the task, so outlook-tasks-list.ps1
# shows "From email" and a second run for the same email is refused unless -Force.
# Usage: .\outlook-tasks-from-mail.ps1 -EntryID "00000000..."
# Due:   .\outlook-tasks-from-mail.ps1 -EntryID "00000000..." -When Tomorrow | -DueDate "2026-11-03"
# More:  -Subject "Reply to Dan" -Priority High -Categories "Teaching" -Reminder "2026-11-02 09:00" -AttachEmail

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")

try {
    $null = Connect-Outlook
    $mail = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID
    if ([int]$mail.Class -ne 43) { Stop-WithError "That EntryID is not an email (item class $([int]$mail.Class))." -Hint "Get email EntryIDs from outlook-find.ps1." }
    $mailId = [string]$mail.EntryID
    $messageId = ""; try { $messageId = [string]$mail.PropertyAccessor.GetProperty($script:TaskProp.InternetMessageID) } catch { }

    if (-not $Force) {
        # Match the Message-ID too: moving the email changes its EntryID.
        $existing = @((Get-TaskRecords -Namespace $Namespace -IncludeCompleted).Records | Where-Object { $_.SourceEntryID -eq $mailId -or ($messageId -and $_.SourceMessageID -eq $messageId) })
        if ($existing.Count -gt 0) {
            Stop-WithError "A task for this email already exists (EntryID $($existing[0].EntryID))." -Hint "Use that task, or pass -Force to create another."
        }
    }

    $sender = $mail.SenderName
    $smtp = Get-SenderSmtp $mail
    if ($smtp -and $smtp -ne $sender) { $sender = "$sender <$smtp>" }
    $received = $mail.ReceivedTime

    if (-not $DueDate -and $When -ne "NoDate") {
        $today = (Get-Date).Date
        $toFriday = [int](5 - [int]$today.DayOfWeek); if ($toFriday -lt 0) { $toFriday += 7 }
        $DueDate = switch ($When) {
            "Today" { $today }
            "Tomorrow" { $today.AddDays(1) }
            "ThisWeek" { $today.AddDays($toFriday) }
            "NextWeek" { $today.AddDays($toFriday + 7) }
        }
    }

    $task = $Outlook.CreateItem(3)   # olTaskItem
    $task.Subject = $(if ($Subject) { $Subject } else { [string]$mail.Subject })
    $task.Importance = $(if ($Priority) { $script:TaskPriorityMap[$Priority] } else { [int]$mail.Importance })
    $lines = @(
        "From email: $sender",
        "Received: $(Format-MailDate $received)",
        "Subject: $($mail.Subject)",
        "Open the email: outlook:$mailId",
        "",
        (Format-Preview ([string]$mail.Body) 500)
    )
    $task.Body = ($lines -join "`r`n")
    if ($DueDate) {
        # Start today, or on the due date when that is already past (start must not follow due).
        $startDay = (Get-Date).Date; if ($DueDate.Date -lt $startDay) { $startDay = $DueDate.Date }
        $task.StartDate = $startDay; $task.DueDate = $DueDate
    }
    if ($Reminder) { $task.ReminderSet = $true; $task.ReminderTime = $Reminder }
    if ($Categories) { $task.Categories = $Categories }
    if ($AttachEmail) { [void]$task.Attachments.Add($mail, 5) }   # olEmbeddeditem: a copy; the email stays where it is
    $task.PropertyAccessor.SetProperty($script:TaskProp.SourceEntryID, $mailId)
    if ($messageId) { $task.PropertyAccessor.SetProperty($script:TaskProp.SourceMessageID, $messageId) }
    $task.Save()

    $changes = [ordered]@{ "From email" = $sender; Received = Format-MailDate $received }
    if ($AttachEmail) { $changes["Email attached"] = $true }
    Write-TaskResult -Title "TASK CREATED FROM EMAIL" -Task $task -Changes $changes -Json:$Json -Action "from-mail"
}
catch {
    Stop-WithError $_
}
