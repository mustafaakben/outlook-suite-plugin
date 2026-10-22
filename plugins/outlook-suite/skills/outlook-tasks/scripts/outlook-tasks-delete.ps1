param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [string]$Subject = "",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Task Delete: move a task to Deleted Items.
# Irreversible action: previews unless -Confirm (or confirmation for "delete" is turned off;
# see the outlook-mail SKILL.md "Confirmation").
# Usage: .\outlook-tasks-delete.ps1 -EntryID "00000000..." -Confirm (preferred)
# Usage: .\outlook-tasks-delete.ps1 -Index 1 -Confirm (position in outlook-tasks-list.ps1 order)
# By subject: .\outlook-tasks-delete.ps1 -Subject "Old task" -Confirm (exact, must be unique)

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")

try {
    $null = Connect-Outlook
    $task = Resolve-TaskItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Subject $Subject -IncludeCompleted:([bool]$Subject)
    if ((Confirm-TaskAction -Action "delete" -Title "DELETE TASK (move to Deleted Items)" -Task $task -Confirmed $Confirm -What "the delete" -Json:$Json) -ne $true) { if ($script:TaskPreviewJson) { Write-Json $script:TaskPreviewJson }; exit 0 }

    $rec = ConvertTo-TaskRecord $task
    $task.Delete()
    if ($Json) {
        $o = ConvertTo-TaskJson $rec
        $o.Action = "delete"; $o.Result = "deleted"
        Write-Json $o
        exit 0
    }
    Write-TaskRecord $rec "=== TASK DELETED (moved to Deleted Items) ==="
    Write-Host "  (The EntryID above is no longer valid.)" -ForegroundColor DarkGray
}
catch {
    Stop-WithError $_
}
