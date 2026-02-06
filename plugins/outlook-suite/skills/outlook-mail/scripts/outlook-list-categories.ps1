# Outlook Category Lister
# Usage: .\outlook-list-categories.ps1

$ColorNames = @{
    0 = "None"; 1 = "Red"; 2 = "Orange"; 3 = "Peach"; 4 = "Yellow"
    5 = "Green"; 6 = "Teal"; 7 = "Olive"; 8 = "Blue"; 9 = "Purple"
    10 = "Maroon"; 11 = "Steel"; 12 = "DarkSteel"; 13 = "Gray"; 14 = "DarkGray"
    15 = "Black"; 16 = "DarkRed"; 17 = "DarkOrange"; 18 = "DarkPeach"; 19 = "DarkYellow"
    20 = "DarkGreen"; 21 = "DarkTeal"; 22 = "DarkOlive"; 23 = "DarkBlue"; 24 = "DarkPurple"
    25 = "DarkMaroon"
}

$Outlook = $null
$Namespace = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")

    Write-Host "`n=== OUTLOOK CATEGORIES ===" -ForegroundColor Cyan
    Write-Host ""

    $count = 0
    foreach ($category in $Namespace.Categories) {
        $count++
        $colorIndex = $category.Color
        $colorName = if ($ColorNames.ContainsKey($colorIndex)) { $ColorNames[$colorIndex] } else { "Unknown" }

        Write-Host "$count. $($category.Name)" -ForegroundColor Yellow
        Write-Host "      Color: $colorName (index: $colorIndex)" -ForegroundColor Gray
    }

    if ($count -eq 0) {
        Write-Host "No categories defined." -ForegroundColor Gray
        Write-Host "Use outlook-create-category.ps1 to create one." -ForegroundColor Gray
    } else {
        Write-Host "`n--- Total: $count categories ---" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Namespace) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
