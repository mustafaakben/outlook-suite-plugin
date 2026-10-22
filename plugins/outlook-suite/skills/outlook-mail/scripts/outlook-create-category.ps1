param(
    [Parameter(Mandatory = $true)]
    [string]$Name,
    [string]$Color = "Red",
    [switch]$Json
)

# Outlook Create Category: add a category to the master list.
# -Color: a name (Red, Green, Blue, DarkPurple, ...) or an index 0-25.
# Change an existing one with outlook-update-category.ps1.
# Usage: .\outlook-create-category.ps1 -Name "@Action" -Color Red

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $idx = Resolve-CategoryColor $Color
    $existing = Get-OutlookCategory $Namespace $Name
    if ($existing) { Stop-WithError "Category '$($existing.Name)' already exists (color $($script:CategoryColors[[int]$existing.Color]))." "Change it with outlook-update-category.ps1 -Name ""$Name"" -Color <color>." }
    $null = $Namespace.Categories.Add($Name, $idx)
    $o = [ordered]@{ Result = "created"; Name = $Name; Color = $script:CategoryColors[$idx] }
    if ($Json) { Write-Json $o; exit 0 }
    Write-Host "=== CATEGORY CREATED ===" -ForegroundColor Green
    Write-Host "  Name: $Name   Color: $($o.Color)"
}
catch {
    Stop-WithError $_
}
