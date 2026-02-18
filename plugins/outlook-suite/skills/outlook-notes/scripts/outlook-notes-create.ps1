param(
    [Parameter(Mandatory=$true)]
    [string]$Body,

    [ValidateSet("Blue", "Green", "Pink", "Yellow", "White")]
    [string]$Color = "Yellow",
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Note Create Script
# Usage: .\outlook-notes-create.ps1 -Body "Remember to call John"
# With color: .\outlook-notes-create.ps1 -Body "Important meeting notes" -Color Pink

# Note color constants
# 0 = Blue
# 1 = Green
# 2 = Pink
# 3 = Yellow (default)
# 4 = White

$colorMap = @{
    "Blue" = 0
    "Green" = 1
    "Pink" = 2
    "Yellow" = 3
    "White" = 4
}

$colorDisplay = @{
    "Blue" = "Cyan"
    "Green" = "Green"
    "Pink" = "Magenta"
    "Yellow" = "Yellow"
    "White" = "White"
}

$Outlook = $null
$note = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    # Create note item (olNoteItem = 5)
    $note = $Outlook.CreateItem(5)

    $note.Body = $Body
    $note.Color = $colorMap[$Color]

    $note.Save()

    # Get preview
    $preview = $Body -replace "`r`n", " " -replace "`n", " "
    if ($preview.Length -gt 60) {
        $preview = $preview.Substring(0, 57) + "..."
    }
    if ($StripLinks) { $preview = Strip-Links $preview }

    Write-Host "`n=== NOTE CREATED ===" -ForegroundColor Green
    Write-Host "Color: $Color" -ForegroundColor $colorDisplay[$Color]
    Write-Host "Content: $preview" -ForegroundColor Gray

    # Show full content if multiline
    $lines = $Body -split "`n"
    if ($lines.Count -gt 1) {
        $displayBody = $Body
        if ($StripLinks) { $displayBody = Strip-Links $displayBody }
        Write-Host "`nFull content ($($lines.Count) lines):" -ForegroundColor Cyan
        Write-Host $displayBody -ForegroundColor $colorDisplay[$Color]
    }

    Write-Host "`nNote saved to Notes folder." -ForegroundColor Cyan
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($note) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($note) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
