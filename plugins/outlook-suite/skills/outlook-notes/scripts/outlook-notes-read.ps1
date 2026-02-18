param(
    [string]$EntryID = "",
    [int]$Index = 0,

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

# Outlook Note Read Script
# Read full content of a note
# Usage: .\outlook-notes-read.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-notes-read.ps1 -Index 1 (fallback)
# With color filter: .\outlook-notes-read.ps1 -Index 1 -Color Yellow

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
    0 = "Cyan"
    1 = "Green"
    2 = "Magenta"
    3 = "Yellow"
    4 = "White"
}

function Test-IsOutlookNoteItem {
    param([object]$Item)

    if (-not $Item) {
        return $false
    }

    try {
        $messageClass = $Item.MessageClass
        if ($messageClass -and $messageClass -like "IPM.StickyNote*") {
            return $true
        }
    } catch {}

    try {
        return ([int]$Item.Class -eq 44)  # olNote
    } catch {
        return $false
    }
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
    $targetNote = $null

    if ($EntryID) {
        $resolvedItem = $Namespace.GetItemFromID($EntryID)
        if (-not $resolvedItem) {
            Write-Host "ERROR: Note not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-notes-list.ps1 to find notes and get EntryIDs." -ForegroundColor Gray
        } elseif (-not (Test-IsOutlookNoteItem -Item $resolvedItem)) {
            Write-Host "ERROR: Provided EntryID does not resolve to an Outlook note." -ForegroundColor Red
            Write-Host "Tip: Use outlook-notes-list.ps1 to get EntryIDs from notes only." -ForegroundColor Gray
        } else {
            $targetNote = $resolvedItem
        }
    } elseif ($Index -gt 0) {
        $Notes = $Namespace.GetDefaultFolder(12)  # olFolderNotes = 12
        $items = $Notes.Items
        $items.Sort("[LastModificationTime]", $true)

        # Build filtered list
        $filteredNotes = @()
        foreach ($note in $items) {
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

            $filteredNotes += $note
        }

        # Find the note at the specified index
        if ($Index -lt 1 -or $Index -gt $filteredNotes.Count) {
            Write-Host "`nNote at index $Index not found." -ForegroundColor Red
            if ($Color -ne "All" -or $Search) {
                Write-Host "Use outlook-notes-list.ps1 with same filters to see available notes." -ForegroundColor Gray
            } else {
                Write-Host "Use outlook-notes-list.ps1 to see available notes." -ForegroundColor Gray
            }
        } else {
            $targetNote = $filteredNotes[$Index - 1]
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to read a note." -ForegroundColor Red
        Write-Host "Usage: .\outlook-notes-read.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if ($targetNote) {
        # Get note details
        $noteColor = $colorNames[$targetNote.Color]
        $displayColor = $colorDisplay[$targetNote.Color]
        $body = if ($targetNote.Body) { $targetNote.Body } else { "(Empty note)" }
        $created = $targetNote.CreationTime
        $modified = $targetNote.LastModificationTime

        Write-Host "`n=== NOTE CONTENT ===" -ForegroundColor Cyan
        Write-Host "Color: $noteColor" -ForegroundColor $displayColor
        Write-Host "Created: $($created.ToString('g'))" -ForegroundColor Gray
        Write-Host "Modified: $($modified.ToString('g'))" -ForegroundColor Gray

        if ($StripLinks) { $body = Strip-Links $body }
        Write-Host "`n--- Content ---" -ForegroundColor Cyan
        Write-Host $body -ForegroundColor $displayColor

        # Show stats
        $charCount = $body.Length
        $lineCount = ($body -split "`n").Count
        Write-Host "`n($charCount characters, $lineCount lines)" -ForegroundColor Gray
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
