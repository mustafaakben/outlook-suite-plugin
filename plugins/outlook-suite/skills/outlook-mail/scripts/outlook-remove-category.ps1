param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [string[]]$Category = @(),
    [switch]$All,
    [int]$Days = 7,
    [switch]$Json
)

# Outlook Remove Category: remove one or more categories from an email, or all (-All).
# Usage: .\outlook-remove-category.ps1 -EntryID "00000000..." -Category "@Action"
# Usage: .\outlook-remove-category.ps1 -EntryID "00000000..." -All

. (Join-Path $PSScriptRoot "_common.ps1")

if (-not $All -and $Category.Count -eq 0) { Stop-WithError "Give -Category <name> or -All." }

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $current = @(([string]$email.Categories).Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() })
    if ($All) { $removed = $current; $keep = @() }
    else {
        $drop = @($Category | ForEach-Object { $_.Split(',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $removed = @($current | Where-Object { $drop -contains $_ })
        $keep = @($current | Where-Object { $drop -notcontains $_ })
    }
    if ($removed.Count -gt 0) { $email.Categories = $keep -join ", "; $email.Save() }
    $changes = [ordered]@{ Removed = $(if ($removed.Count) { $removed -join ", " } else { "(none, not assigned)" }); Categories = $(if ($keep.Count) { $keep -join ", " } else { "(none)" }) }
    Write-ActionResult -Title "CATEGORY REMOVED" -Item $email -Changes $changes -Json:$Json -Action "remove-category"
}
catch {
    Stop-WithError $_
}
