---
name: outlook-notes
description: "Manage Outlook sticky notes via PowerShell: list, search the full note text, filter by color or category, read, create and delete notes. Fast table queries, -Json output, confirmation for delete. Auto-starts Outlook if needed."
---

# Outlook Notes

Manage Microsoft Outlook sticky notes with PowerShell COM scripts. All scripts share the outlook-suite helpers in `../outlook-mail/scripts/_common.ps1` (loaded through `scripts/_common.ps1`): same connection handling, `-Json`, errors as `Error: ...` with exit code 1, and the same confirmation gate as the mail scripts.

## Prerequisites

- Classic Microsoft Outlook for Windows (scripts start it with its window if it is not running).
- Windows PowerShell 5.1 or PowerShell 7.

## Fast path (warm host)

```powershell
& "../outlook-mail/scripts/outlook-call.ps1" notes-list -Search "budget"
& "../outlook-mail/scripts/outlook-call.ps1" notes-read -EntryID "00000000..."
```

## Links

Scripts that show note text take `-StripLinks` (default `$true`): links appear as `[URL-001]`, and `outlook-notes-read.ps1` lists the full URLs at the end. Pass `-StripLinks $false` (or `false` / `0`) to keep links inline.

## Choosing a note

`-EntryID` (preferred, from `outlook-notes-list.ps1`), or `-Index N`: the N-th note in `outlook-notes-list.ps1` order with the same `-Color`, `-Search` and `-Category` filters. Index positions shift as notes change.

## outlook-notes-list.ps1

Notes, most recently changed first, from one fast table query. Each shows color, a one-line preview, EntryID, categories and when it was last changed.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Search` | string | — | The note text contains this (whole note, plain substring, any case) |
| `-Color` | string | All | All, Blue, Green, Pink, Yellow, White |
| `-Category` | string | — | Has this category |
| `-Limit` | int | 20 | Maximum notes shown (the total matching is always reported) |
| `-StripLinks` | bool | $true | See Links |
| `-Json` | switch | false | `{ folder, total, count, query, items[] }`; each item has `Title` (first line) and `Preview` (first 255 characters) |

```powershell
& "./scripts/outlook-notes-list.ps1"
& "./scripts/outlook-notes-list.ps1" -Search "meeting" -Color Pink
& "./scripts/outlook-notes-list.ps1" -Category "Research" -Json
```

## outlook-notes-read.ps1

The full text of one note, with color, categories, created and modified dates. `-Json` returns the note with `Body` and `Links`. Targeting: `-EntryID`, or `-Index` with the list filters.

```powershell
& "./scripts/outlook-notes-read.ps1" -EntryID "00000000..."
& "./scripts/outlook-notes-read.ps1" -Index 1 -Search "budget"
```

## outlook-notes-create.ps1

Create a note (a local item; nothing is sent). The first line of the text becomes its title.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Body` | string | **Required** | Note text (multi-line allowed) |
| `-Color` | string | Yellow | Blue, Green, Pink, Yellow, White |
| `-Categories` | string | — | e.g. `"Research"` |
| `-Json` | switch | false | The new note as JSON (with its EntryID) |

```powershell
& "./scripts/outlook-notes-create.ps1" -Body "Remember to call John"
& "./scripts/outlook-notes-create.ps1" -Body "Idea: ...`nSecond line" -Color Pink -Categories "Research"
```

## outlook-notes-delete.ps1

Move a note to Deleted Items. Previews unless `-Confirm` (follows the `delete` confirmation setting; see the outlook-mail SKILL.md). Targeting as in read. `-Json` available.

```powershell
& "./scripts/outlook-notes-delete.ps1" -EntryID "00000000..."            # preview
& "./scripts/outlook-notes-delete.ps1" -EntryID "00000000..." -Confirm
```

## Reference: note colors

| Value | Color |
|-------|-------|
| 0 | Blue (also notes with no color set) |
| 1 | Green |
| 2 | Pink |
| 3 | Yellow |
| 4 | White |

Tests: `tests/Test-Notes.ps1` (live, on `[suite-test]` notes; prints PASS/FAIL only).
