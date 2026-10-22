param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [Parameter(Mandatory = $true)]
    [string[]]$Category,
    [int]$Days = 7,
    [switch]$Json
)

# Outlook Assign Category: add one or more categories to an email (existing ones are kept).
# Categories not in your master list still work; create them with outlook-create-category.ps1
# to give them a color.
# Usage: .\outlook-assign-category.ps1 -EntryID "00000000..." -Category "@Action"
# Usage: .\outlook-assign-category.ps1 -EntryID "00000000..." -Category "@Action","Project X"

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $current = @(([string]$email.Categories).Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() })
    $wanted = @($Category | ForEach-Object { $_.Split(',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $added = @($wanted | Where-Object { $current -notcontains $_ })
    if ($added.Count -gt 0) {
        $email.Categories = (@($current) + $added) -join ", "
        $email.Save()
    }
    $master = @(foreach ($c in $Namespace.Categories) { $c.Name })
    $notInMaster = @($wanted | Where-Object { $master -notcontains $_ })
    $changes = [ordered]@{ Added = $(if ($added.Count) { $added -join ", " } else { "(none, already assigned)" }); Categories = $email.Categories }
    if ($notInMaster.Count -gt 0) { $changes.Note = "Not in the master category list (no color): $($notInMaster -join ', ')" }
    Write-ActionResult -Title "CATEGORY ASSIGNED" -Item $email -Changes $changes -Json:$Json -Action "assign-category"
}
catch {
    Stop-WithError $_
}
