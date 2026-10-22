param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [string]$Subject = "",
    [switch]$Json
)

# Outlook Task Complete: mark a task complete. A flagged email's EntryID (from
# outlook-tasks-list.ps1 -Flagged) marks its follow-up flag complete instead.
# Usage: .\outlook-tasks-complete.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-tasks-complete.ps1 -Index 1 (position in outlook-tasks-list.ps1 order)
# By subject: .\outlook-tasks-complete.ps1 -Subject "Finish report" (exact, must be unique)

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")

try {
    $null = Connect-Outlook
    $item = Resolve-TaskItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Subject $Subject -AllowMail

    if ([int]$item.Class -eq 43) {
        # Flagged email: complete the flag (olFlagComplete). The email itself is not changed otherwise.
        if ([int]$item.FlagStatus -eq 1) {
            if ($Json) { Write-Json ([ordered]@{ Action = "complete"; Result = "already complete"; EntryID = $item.EntryID; Type = "Mail"; Subject = $item.Subject }); exit 0 }
            Write-Host "Flag is already complete: $($item.Subject)" -ForegroundColor Yellow; exit 0
        }
        $item.FlagStatus = 1
        $item.Save()
        if ($Json) { Write-Json ([ordered]@{ Action = "complete"; Result = "flag completed"; EntryID = $item.EntryID; Type = "Mail"; Subject = $item.Subject }); exit 0 }
        Write-Host "=== EMAIL FLAG COMPLETED ===" -ForegroundColor Green
        Write-Host "  EntryID: $($item.EntryID)"
        Write-Host "  Subject: $($item.Subject)" -ForegroundColor Yellow
        exit 0
    }

    if ($item.Complete) {
        Write-TaskResult -Title "TASK ALREADY COMPLETED" -Task $item -Json:$Json -Action "complete"
        exit 0
    }
    $item.MarkComplete()
    Write-TaskResult -Title "TASK COMPLETED" -Task $item -Json:$Json -Action "complete"
}
catch {
    Stop-WithError $_
}
