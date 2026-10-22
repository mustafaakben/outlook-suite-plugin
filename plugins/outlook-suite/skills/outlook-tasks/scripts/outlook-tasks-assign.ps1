param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [string]$AssignTo,

    [switch]$KeepCopy,
    [switch]$IncludeCompleted,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Task Assign: send a task request to another person (needs Exchange / Microsoft 365).
# Outward-facing: this SENDS an email. Previews unless -Confirm (or confirmation for "send"
# is turned off; see the outlook-mail SKILL.md "Confirmation").
# Usage: .\outlook-tasks-assign.ps1 -EntryID "00000000..." -AssignTo "colleague@example.com" -Confirm
# Keep copy: .\outlook-tasks-assign.ps1 -EntryID "00000000..." -AssignTo "colleague@example.com" -KeepCopy -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")

try {
    $null = Connect-Outlook
    $task = Resolve-TaskItem -Namespace $Namespace -EntryID $EntryID -Index $Index -IncludeCompleted:$IncludeCompleted

    # Resolve the person before anything else, so a typo fails without touching the task.
    $who = $Namespace.CreateRecipient($AssignTo)
    if (-not $who.Resolve()) { Stop-WithError "Could not resolve '$AssignTo'." -Hint "Use a full email address, or find the person with the contacts skill." }
    $whoName = $who.Name
    $whoSmtp = ""
    try { $ex = $who.AddressEntry.GetExchangeUser(); if ($ex) { $whoSmtp = $ex.PrimarySmtpAddress } } catch { }
    if (-not $whoSmtp) { try { $whoSmtp = $who.AddressEntry.Address } catch { } }
    $whoLabel = if ($whoSmtp -and $whoSmtp -ne $whoName) { "$whoName <$whoSmtp>" } else { $whoName }

    $details = [ordered]@{ "Assign to" = $whoLabel; "Keep copy" = [bool]$KeepCopy }
    if ($task.Delegator) { $details["Already assigned by"] = $task.Delegator }
    if ((Confirm-TaskAction -Action "send" -Title "ASSIGN TASK (sends a task request email)" -Task $task -Confirmed $Confirm -Details $details -What "the task request" -Json:$Json) -ne $true) { if ($script:TaskPreviewJson) { Write-Json $script:TaskPreviewJson }; exit 0 }

    $rec = ConvertTo-TaskRecord $task
    $request = $task.Assign()
    [void]$request.Recipients.Add($(if ($whoSmtp) { $whoSmtp } else { $AssignTo }))   # the person resolved above
    if (-not $request.Recipients.ResolveAll()) { Stop-WithError "Could not resolve '$AssignTo' on the task request; nothing was sent." }
    if ($KeepCopy) { $request.KeepAfterAssignment = $true }
    $request.Send()

    $changes = [ordered]@{ "Assigned to" = $whoLabel; "Keep copy" = [bool]$KeepCopy }
    Write-TaskResult -Title "TASK ASSIGNED (request sent)" -Record $rec -Changes $changes -Json:$Json -Action "assign"
}
catch {
    if ($_.Exception.Message -like "*Operation aborted*") { Stop-WithError "Task assignment was cancelled." }
    Stop-WithError $_ -Hint "Task assignment requires Exchange or Microsoft 365."
}
