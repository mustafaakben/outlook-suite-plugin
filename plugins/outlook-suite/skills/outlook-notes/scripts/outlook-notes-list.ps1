param(
    [ValidateRange(1, 2147483647)]
    [int]$Limit = 20,
    [ValidateSet("All", "Blue", "Green", "Pink", "Yellow", "White")]
    [string]$Color = "All",
    [string]$Search = "",
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Notes List Script
# Usage: .\outlook-notes-list.ps1
# With limit: .\outlook-notes-list.ps1 -Limit 50
# Filter by color: .\outlook-notes-list.ps1 -Color Yellow
# Search: .\outlook-notes-list.ps1 -Search "meeting"

# Note color constants
# 0 = Blue
# 1 = Green
# 2 = Pink
# 3 = Yellow
# 4 = White

$colorMap = @{
    "Blue" = 0
    "Green" = 1
    "Pink" = 2
    "Yellow" = 3
    "White" = 4
}

$colorNames = @{
    0 = "Blue"
    1 = "Green"
    2 = "Pink"
    3 = "Yellow"
    4 = "White"
}

$colorDisplay = @{
    0 = "Cyan"       # Blue note -> Cyan in console
    1 = "Green"
    2 = "Magenta"    # Pink note -> Magenta in console
    3 = "Yellow"
    4 = "White"
}

$Outlook = $null
$Namespace = $null
$Notes = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")
    $Notes = $Namespace.GetDefaultFolder(12)  # olFolderNotes = 12

    $items = $Notes.Items
    $items.Sort("[LastModificationTime]", $true)

    Write-Host "`n=== OUTLOOK NOTES ===" -ForegroundColor Cyan

    if ($Color -ne "All") {
        Write-Host "(Color: $Color)" -ForegroundColor Gray
    }

    if ($Search) {
        Write-Host "(Search: '$Search')" -ForegroundColor Gray
    }

    Write-Host ""

    $count = 0
    $displayed = 0

    foreach ($note in $items) {
        $count++

        # Filter by color
        if ($Color -ne "All") {
            if ($note.Color -ne $colorMap[$Color]) {
                continue
            }
        }

        # Filter by search
        if ($Search) {
            $searchLower = $Search.ToLower()
            $body = if ($note.Body) { $note.Body.ToLower() } else { "" }
            if (-not $body.Contains($searchLower)) {
                continue
            }
        }

        $displayed++
        if ($displayed -gt $Limit) { break }

        # Get note preview (first line or first 60 chars)
        $body = if ($note.Body) { $note.Body } else { "(Empty note)" }
        $preview = $body -replace "`r`n", " " -replace "`n", " "
        if ($preview.Length -gt 60) {
            $preview = $preview.Substring(0, 57) + "..."
        }
        if ($StripLinks) { $preview = Strip-Links $preview }

        # Get note color
        $noteColor = $colorNames[$note.Color]
        $displayColor = $colorDisplay[$note.Color]

        # Display note
        Write-Host "$displayed. " -NoNewline
        Write-Host "[$noteColor] " -NoNewline -ForegroundColor $displayColor
        Write-Host $preview -ForegroundColor $displayColor
        Write-Host "      EntryID: $($note.EntryID)" -ForegroundColor DarkGray

        # Show modification time
        $modified = $note.LastModificationTime
        $timeAgo = ""
        $diff = (Get-Date) - $modified
        if ($diff.TotalDays -ge 1) {
            $timeAgo = "$([math]::Floor($diff.TotalDays)) days ago"
        } elseif ($diff.TotalHours -ge 1) {
            $timeAgo = "$([math]::Floor($diff.TotalHours)) hours ago"
        } else {
            $timeAgo = "$([math]::Floor($diff.TotalMinutes)) minutes ago"
        }

        Write-Host "      Modified: $($modified.ToString('g')) ($timeAgo)" -ForegroundColor Gray
        Write-Host ""
    }

    if ($displayed -eq 0) {
        if ($Search -or $Color -ne "All") {
            Write-Host "No notes found matching criteria." -ForegroundColor Gray
        } else {
            Write-Host "No notes found." -ForegroundColor Gray
        }
    } else {
        Write-Host "--- Showing $displayed notes ---" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Notes) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Notes) | Out-Null } catch {}
    }
    if ($Namespace) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
