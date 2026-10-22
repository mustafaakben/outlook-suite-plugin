param(
    [Parameter(Mandatory = $true)]
    [string]$Name,
    [string]$NewName = "",
    [string]$Color = "",
    [ValidateSet("", "None", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12")]
    [string]$Shortcut = "",
    [switch]$Json
)

# Outlook Update Category: rename a master category, change its color or Ctrl+F<n> shortcut.
# Renaming changes the master list only; emails already tagged keep the old name as text
# (outlook-bulk.ps1 can re-tag them).
# Usage: .\outlook-update-category.ps1 -Name "@Action" -Color DarkRed -Shortcut F2
# Usage: .\outlook-update-category.ps1 -Name "Old" -NewName "New"

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    if (-not ($NewName -or $Color -or $Shortcut)) { Stop-WithError "Give -NewName, -Color and/or -Shortcut." }
    $null = Connect-Outlook
    $cat = Get-OutlookCategory $Namespace $Name
    if (-not $cat) { Stop-WithError "Category '$Name' not found. See outlook-list-categories.ps1." }
    $changes = [ordered]@{}
    if ($Color) { $cat.Color = Resolve-CategoryColor $Color; $changes.Color = $script:CategoryColors[[int]$cat.Color] }
    if ($Shortcut) {
        $key = if ($Shortcut -eq "None") { 0 } else { [int]$Shortcut.Substring(1) - 1 }
        $cat.ShortcutKey = $key; $changes.Shortcut = $(if ($key) { "Ctrl+$Shortcut" } else { "none" })
    }
    if ($NewName) {
        if ((Get-OutlookCategory $Namespace $NewName) -and $NewName -ine $Name) { Stop-WithError "A category named '$NewName' already exists." }
        $cat.Name = $NewName; $changes.Renamed = "$Name -> $NewName"
    }
    $o = [ordered]@{ Result = "updated"; Name = $cat.Name } + $changes
    if ($Json) { Write-Json $o; exit 0 }
    Write-Host "=== CATEGORY UPDATED ===" -ForegroundColor Green
    foreach ($k in $o.Keys) { if ($k -ne "Result") { Write-Host "  ${k}: $($o[$k])" } }
}
catch {
    Stop-WithError $_
}
