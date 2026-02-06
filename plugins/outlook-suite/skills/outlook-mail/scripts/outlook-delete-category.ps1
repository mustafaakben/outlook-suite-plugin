param(
    [Parameter(Mandatory=$true)]
    [string]$Name
)

# Outlook Category Deleter
# Usage: .\outlook-delete-category.ps1 -Name "Category Name"
# Deletes a category from the master category list

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

    # Find the category
    $targetIndex = -1
    $index = 1
    foreach ($cat in $Namespace.Categories) {
        if ($cat.Name -eq $Name) {
            $targetIndex = $index
            break
        }
        $index++
    }

    if ($targetIndex -eq -1) {
        Write-Host "`nCategory '$Name' not found." -ForegroundColor Red
        Write-Host "`nAvailable categories:" -ForegroundColor Yellow
        foreach ($cat in $Namespace.Categories) {
            Write-Host "  - $($cat.Name)" -ForegroundColor Gray
        }
    } else {
        # Delete the category
        $Namespace.Categories.Remove($targetIndex)

        Write-Host "`n=== CATEGORY DELETED ===" -ForegroundColor Green
        Write-Host "Name: $Name" -ForegroundColor Yellow
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
