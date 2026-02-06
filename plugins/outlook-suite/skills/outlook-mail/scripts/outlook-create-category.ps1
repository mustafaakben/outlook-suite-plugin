param(
    [Parameter(Mandatory=$true)]
    [string]$Name,

    [int]$Color = 1
)

# Outlook Category Creator
# Usage: .\outlook-create-category.ps1 -Name "Project Alpha" -Color 5
# Colors: 0=None, 1=Red, 2=Orange, 3=Peach, 4=Yellow, 5=Green, 6=Teal, 7=Olive, 8=Blue, 9=Purple, 10=Maroon

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

    # Validate color index
    if ($Color -lt 0 -or $Color -gt 25) {
        Write-Host "`nInvalid color index: $Color. Must be 0-25." -ForegroundColor Red
        Write-Host "Use -Color 0 for None, 1 for Red, 5 for Green, 8 for Blue, etc." -ForegroundColor Gray
    } else {
        # Check if category already exists
        $existing = $null
        foreach ($cat in $Namespace.Categories) {
            if ($cat.Name -eq $Name) {
                $existing = $cat
                break
            }
        }

        if ($existing) {
            $colorName = if ($ColorNames.ContainsKey($existing.Color)) { $ColorNames[$existing.Color] } else { "Unknown" }
            Write-Host "`nCategory '$Name' already exists." -ForegroundColor Red
            Write-Host "Current color: $colorName (index: $($existing.Color))" -ForegroundColor Gray
        } else {
            # Create the category
            $null = $Namespace.Categories.Add($Name, $Color)

            $colorName = if ($ColorNames.ContainsKey($Color)) { $ColorNames[$Color] } else { "Unknown" }

            Write-Host "`n=== CATEGORY CREATED ===" -ForegroundColor Green
            Write-Host "Name: $Name" -ForegroundColor Yellow
            Write-Host "Color: $colorName (index: $Color)" -ForegroundColor Gray
        }
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
