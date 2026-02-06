---
name: outlook-notes
description: "Manage Outlook notes via PowerShell: list, create, read, and delete sticky notes with color coding and search. Auto-starts Outlook if needed."
---

# Outlook Notes

Manage Microsoft Outlook sticky notes using PowerShell COM objects.

## Prerequisites

- Microsoft Outlook installed (scripts auto-start Outlook if it is not already running)
- PowerShell 5.1+ (included in Windows 10/11)

## List Notes

### outlook-notes-list.ps1

Browse notes with optional color filtering and search. Returns EntryIDs for use with action scripts.

**Required Parameters:** None (lists all notes by default)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Limit` | int | 20 | Maximum notes to display |
| `-Color` | string | All | Filter by color: All, Blue, Green, Pink, Yellow, White |
| `-Search` | string | — | Search note content (case-insensitive) |

```powershell
# List all notes
& "./scripts/outlook-notes-list.ps1"

# List yellow notes only
& "./scripts/outlook-notes-list.ps1" -Color Yellow

# Search notes
& "./scripts/outlook-notes-list.ps1" -Search "meeting"

# Combined filter
& "./scripts/outlook-notes-list.ps1" -Color Pink -Search "important"
```

## Create Note

### outlook-notes-create.ps1

Create a new sticky note with optional color.

**Required Parameters:** `-Body`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Body` | string | **Required** | Note content text |
| `-Color` | string | Yellow | Note color: Blue, Green, Pink, Yellow, White |

```powershell
# Create a note (default yellow)
& "./scripts/outlook-notes-create.ps1" -Body "Remember to call John"

# Create a colored note
& "./scripts/outlook-notes-create.ps1" -Body "Important meeting notes" -Color Pink

# Create a multiline note
& "./scripts/outlook-notes-create.ps1" -Body "Line 1`nLine 2`nLine 3" -Color Blue
```

## Read Note

### outlook-notes-read.ps1

Read the full content of a note. Uses EntryID (preferred) or Index as targeting method.

**Required Parameters:** `-EntryID` or `-Index`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Note's unique EntryID (preferred, from list) |
| `-Index` | int | 0 | Note position number (fallback) |
| `-Color` | string | All | Filter by color before applying Index (ignored with EntryID) |
| `-Search` | string | — | Search filter before applying Index (ignored with EntryID) |

```powershell
# Read by EntryID (preferred)
& "./scripts/outlook-notes-read.ps1" -EntryID "00000000..."

# Read by index
& "./scripts/outlook-notes-read.ps1" -Index 1

# Read with color filter
& "./scripts/outlook-notes-read.ps1" -Index 1 -Color Yellow
```

## Delete Note

### outlook-notes-delete.ps1

Delete a note from Outlook. Uses EntryID (preferred) or Index as targeting method.

**Required Parameters:** `-EntryID` or `-Index`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Note's unique EntryID (preferred) |
| `-Index` | int | 0 | Note position number (fallback) |
| `-Color` | string | All | Filter by color before applying Index (ignored with EntryID) |
| `-Search` | string | — | Search filter before applying Index (ignored with EntryID) |

```powershell
# Delete by EntryID (preferred)
& "./scripts/outlook-notes-delete.ps1" -EntryID "00000000..."

# Delete by index
& "./scripts/outlook-notes-delete.ps1" -Index 1

# Delete with color filter
& "./scripts/outlook-notes-delete.ps1" -Index 1 -Color Yellow
```

## Reference: Note Colors

| Index | Color |
|-------|-------|
| 0 | Blue |
| 1 | Green |
| 2 | Pink |
| 3 | Yellow |
| 4 | White |
