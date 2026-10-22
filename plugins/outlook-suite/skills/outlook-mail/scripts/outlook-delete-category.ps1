param(
    [Parameter(Mandatory = $true)]
    [string]$Name,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Delete Category: remove a category from the master list.
# Emails that use it keep the name as text but lose the color. Irreversible: previews unless -Confirm.
# Usage: .\outlook-delete-category.ps1 -Name "Old Project" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $idx = 0; $found = $null; $i = 0
    foreach ($c in $Namespace.Categories) { $i++; if ($c.Name -ieq $Name) { $found = $c; $idx = $i; break } }
    if (-not $found) { Stop-WithError "Category '$Name' not found. See outlook-list-categories.ps1." }
    $color = $script:CategoryColors[[int]$found.Color]
    if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "category-delete"))) {
        Write-Host "=== DELETE CATEGORY (PREVIEW) ===" -ForegroundColor Yellow
        Write-Host "  Name: $($found.Name)   Color: $color"
        Approve-Action -Action "category-delete" -Confirmed $false -What "deleting the category" | Out-Null
        exit 0
    }
    $Namespace.Categories.Remove($idx)
    if ($Json) { Write-Json ([ordered]@{ Result = "deleted"; Name = $found.Name; Color = $color }); exit 0 }
    Write-Host "=== CATEGORY DELETED ===" -ForegroundColor Green
    Write-Host "  Name: $Name   Color was: $color"
}
catch {
    Stop-WithError $_
}
