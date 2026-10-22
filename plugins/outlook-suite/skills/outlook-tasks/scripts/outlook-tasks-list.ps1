param(
    [int]$Limit = 20,
    [switch]$IncludeCompleted,
    [ValidateSet("All", "NotStarted", "InProgress", "Completed", "Waiting", "Deferred")]
    [string]$Status = "All",
    [ValidateSet("Any", "Today", "Overdue", "Week", "NoDate")]
    [string]$Due = "Any",
    [switch]$Flagged,
    [string]$Search = "",
    [string]$Category = "",
    [switch]$Json
)

# Outlook Tasks List: open tasks sorted by due date (no due date last), one fast table query.
# Usage: .\outlook-tasks-list.ps1
# Due views: .\outlook-tasks-list.ps1 -Due Today | -Due Overdue | -Due Week (next 7 days) | -Due NoDate
# With flagged email as to-dos: .\outlook-tasks-list.ps1 -Flagged -Due Week
# Filters: -Status InProgress | -Search "report" (subject) | -Category "Work" | -IncludeCompleted
# JSON: .\outlook-tasks-list.ps1 -Due Overdue -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_tasks.ps1")

try {
    $null = Connect-Outlook
    $withCompleted = $IncludeCompleted -or $Status -eq "Completed"
    $res = Get-TaskRecords -Namespace $Namespace -Flagged:$Flagged -IncludeCompleted:$withCompleted
    $matched = @($res.Records | Where-Object {
        ($Status -eq "All" -or $_.Status -eq $Status) -and
        (Test-TaskDue $_ $Due) -and
        (-not $Search -or $_.Subject.IndexOf($Search, [StringComparison]::OrdinalIgnoreCase) -ge 0) -and   # plain substring: -like would treat [ ] as wildcards
        (-not $Category -or ($_.Categories -split '\s*[,;]\s*') -contains $Category)
    })
    $shown = @($matched | Select-Object -First $Limit)
    Add-TaskRecurrenceText $Namespace $shown   # pattern text for the recurring ones shown

    $criteria = @()
    if ($Flagged) { $criteria += "tasks + flagged email" }
    if ($Due -ne "Any") { $criteria += "due: $Due" }
    if ($Status -ne "All") { $criteria += "status: $Status" } elseif (-not $withCompleted) { $criteria += "open only" }
    if ($Search) { $criteria += "subject contains '$Search'" }
    if ($Category) { $criteria += "category: $Category" }

    if ($Json) {
        $items = @(foreach ($r in $shown) { ConvertTo-TaskJson $r })
        Write-Json ([ordered]@{ folder = $res.Folder.FolderPath; total = $matched.Count; count = $items.Count
            query = [ordered]@{ Due = $Due; Status = $Status; Flagged = [bool]$Flagged; Search = $Search; Category = $Category; IncludeCompleted = [bool]$withCompleted }
            items = $items })
        exit 0
    }

    Write-Host "=== OUTLOOK TASKS ===" -ForegroundColor Cyan
    Write-Host "Folder: $($res.Folder.FolderPath)"
    if ($criteria.Count) { Write-Host "Filters: $($criteria -join ' | ')" -ForegroundColor Gray }
    Write-Host ""
    # "#N" only for the unfiltered list: that is the order the action scripts' -Index counts.
    $indexable = $Due -eq "Any" -and -not $Search -and -not $Category -and -not $Flagged -and $Status -eq "All"
    $n = 0
    foreach ($r in $shown) {
        $n++
        Write-TaskRecord $r $(if ($indexable) { "--- #$n ---" } else { "--- result $n ---" })
        Write-Host ""
    }
    if ($n -eq 0) {
        Write-Host "No tasks found." -ForegroundColor Gray
    } else {
        $more = if ($matched.Count -gt $n) { " of $($matched.Count) matching; raise -Limit for more" } else { "" }
        Write-Host "=== Showing $n$more ===" -ForegroundColor Cyan
        Write-Host "Next: outlook-tasks-update.ps1 -EntryID <EntryID> ... | outlook-tasks-complete.ps1 -EntryID <EntryID>" -ForegroundColor DarkGray
    }
}
catch {
    Stop-WithError $_
}
