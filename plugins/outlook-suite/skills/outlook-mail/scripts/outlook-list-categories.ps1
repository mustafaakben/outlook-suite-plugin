param(
    [switch]$Json
)

# Outlook Categories: the master category list (name, color, shortcut key).
# Usage: .\outlook-list-categories.ps1

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $shortcuts = @{ 0 = ""; 1 = "Ctrl+F2"; 2 = "Ctrl+F3"; 3 = "Ctrl+F4"; 4 = "Ctrl+F5"; 5 = "Ctrl+F6"; 6 = "Ctrl+F7"; 7 = "Ctrl+F8"; 8 = "Ctrl+F9"; 9 = "Ctrl+F10"; 10 = "Ctrl+F11"; 11 = "Ctrl+F12" }
    $list = @(foreach ($c in $Namespace.Categories) {
        [ordered]@{ Name = $c.Name; Color = $script:CategoryColors[[int]$c.Color]; ColorIndex = [int]$c.Color; Shortcut = $shortcuts[[int]$c.ShortcutKey] }
    })
    if ($Json) { Write-Json ([ordered]@{ count = $list.Count; categories = $list }); exit 0 }
    Write-Host "=== OUTLOOK CATEGORIES ($($list.Count)) ===" -ForegroundColor Cyan
    foreach ($c in $list) {
        Write-Host "  $($c.Name)  [$($c.Color)]$(if ($c.Shortcut) { "  $($c.Shortcut)" })"
    }
    if ($list.Count -eq 0) { Write-Host "No categories. Create one with outlook-create-category.ps1." -ForegroundColor Gray }
}
catch {
    Stop-WithError $_
}
