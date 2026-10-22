# Helpers for the outlook-notes scripts.
# Dot-source after _common.ps1:  . (Join-Path $PSScriptRoot "_notes.ps1")

$script:NoteColorMap = @{ "Blue" = 0; "Green" = 1; "Pink" = 2; "Yellow" = 3; "White" = 4 }
$script:NoteColorNames = @{ 0 = "Blue"; 1 = "Green"; 2 = "Pink"; 3 = "Yellow"; 4 = "White" }
$script:NoteColorConsole = @{ 0 = "Cyan"; 1 = "Green"; 2 = "Magenta"; 3 = "Yellow"; 4 = "White" }

$script:NoteProp = @{
    Color      = "http://schemas.microsoft.com/mapi/id/{0006200E-0000-0000-C000-000000000046}/8B000003"
    Body       = "urn:schemas:httpmail:textdescription"   # as a column: first 255 chars; in a filter: the whole body
    LongTermID = "http://schemas.microsoft.com/mapi/proptag/0x66700102"
}

# Notes as records, newest change first. -Search matches the whole body (substring, any case)
# inside Outlook's own filter; -Color and -Category filter the same table.
function Get-NoteRecords {
    param($Namespace, [string]$Color = "All", [string]$Search = "", [string]$Category = "")
    $folder = $Namespace.GetDefaultFolder(12)   # olFolderNotes
    $clauses = @()
    # % and _ are LIKE wildcards: Outlook gets the longest stretch without them, and the
    # candidates are then checked for the exact text (opening only those notes).
    $exact = $Search -match '[%_]'
    $likeTerm = $Search
    if ($exact) { $likeTerm = (@($Search -split '[%_]') | Sort-Object Length -Descending | Select-Object -First 1) }
    if ($likeTerm) { $clauses += "`"$($script:NoteProp.Body)`" LIKE '%$(ConvertTo-DaslText $likeTerm)%'" }
    # Categories hold several names ("Research, Teaching"): LIKE narrows, the exact name is checked below.
    if ($Category) { $clauses += "`"urn:schemas-microsoft-com:office:office#Keywords`" LIKE '%$(ConvertTo-DaslText $Category)%'" }
    $table = $null
    if ($clauses.Count) { $table = $folder.GetTable("@SQL=" + ($clauses -join " AND ")) } else { $table = $folder.GetTable() }
    $table.Columns.RemoveAll()
    foreach ($c in @("EntryID", "Subject", "Categories", "CreationTime", "LastModificationTime", $script:NoteProp.Color, $script:NoteProp.Body, $script:NoteProp.LongTermID, "MessageClass")) {
        [void]$table.Columns.Add($c)
    }
    $table.Sort("LastModificationTime", $true)
    $total = $table.GetRowCount()
    $rows = $null
    if ($total -gt 0) { $rows = $table.GetArray($total) }   # direct assignment keeps the 2-D array
    $list = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $total; $i++) {
        if (-not ([string]$rows[$i, 8] -like "IPM.StickyNote*")) { continue }
        # No color property means Outlook's default, which the item reports as 0 (Blue).
        $colorNum = if ($null -ne $rows[$i, 5]) { [int]$rows[$i, 5] } else { 0 }
        $colorName = $script:NoteColorNames[$colorNum]
        if ($Color -ne "All" -and $colorName -ne $Color) { continue }
        if ($Category -and -not (([string]$rows[$i, 2] -split '\s*[,;]\s*') -contains $Category)) { continue }
        if ($exact) {
            $rowId = [string]$rows[$i, 0]   # not inline: Windows PowerShell 5.1 misreads $rows[$i, 0] inside method arguments
            $body = ""; try { $body = [string]$Namespace.GetItemFromID($rowId).Body } catch { }
            if ($body.IndexOf($Search, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        }
        $eid = $rows[$i, 7]; $eid = if ($eid -is [byte[]]) { ConvertTo-HexId $eid } else { [string]$rows[$i, 0] }
        $rec = [ordered]@{
            EntryID    = $eid
            Color      = $colorName
            Title      = [string]$rows[$i, 1]
            Preview    = [string]$rows[$i, 6]
            Categories = [string]$rows[$i, 2]
            Created    = $rows[$i, 3]
            Modified   = $rows[$i, 4]
        }
        $list.Add($rec)
    }
    return , $list
}

# -StripLinks value: accepts $false / "false" / "$false" / 0 / "no" as false. The parameter is
# untyped because pwsh -File passes "$false" as a string, which a [bool] parameter rejects.
function ConvertTo-StripLinks($Value) {
    if ($Value -is [bool]) { return $Value }
    return -not ([string]$Value -match '^\s*(\$?false|0|no|off)\s*$')
}

# One-line preview, links shortened (or kept with -StripLinks $false).
function Format-NotePreview([string]$Text, [bool]$StripLinks = $true, [int]$Length = 60) {
    if (-not $Text) { return "(Empty note)" }
    if ($StripLinks) { return (Format-Preview $Text $Length) }
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $Length) { $t = $t.Substring(0, $Length) + "..." }
    return $t
}

# "3 days ago" / "2 hours ago" / "5 minutes ago".
function Format-Ago($Date) {
    $diff = (Get-Date) - [datetime]$Date
    if ($diff.TotalDays -ge 1) { return "$([math]::Floor($diff.TotalDays)) days ago" }
    if ($diff.TotalHours -ge 1) { return "$([math]::Floor($diff.TotalHours)) hours ago" }
    return "$([math]::Max(0, [math]::Floor($diff.TotalMinutes))) minutes ago"
}

# The note a script targets: -EntryID (preferred) or -Index in outlook-notes-list.ps1 order
# with the same -Color / -Search / -Category filters. Returns the NoteItem.
function Resolve-NoteItem {
    param($Namespace, [string]$EntryID = "", [int]$Index = 0, [string]$Color = "All", [string]$Search = "", [string]$Category = "")
    if ($EntryID) {
        $item = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID
        if ([int]$item.Class -ne 44 -and -not ([string]$item.MessageClass -like "IPM.StickyNote*")) {
            throw "That EntryID is not an Outlook note (item class $([int]$item.Class)). Use outlook-notes-list.ps1 to get note EntryIDs."
        }
        return $item
    }
    if ($Index -gt 0) {
        Write-Host "Note: -Index is volatile (positions shift as notes change). Prefer -EntryID." -ForegroundColor DarkGray
        $recs = Get-NoteRecords -Namespace $Namespace -Color $Color -Search $Search -Category $Category
        if ($Index -gt $recs.Count) { throw "No note at -Index $Index ($($recs.Count) notes match). Use -EntryID from outlook-notes-list.ps1." }
        return (Get-OutlookItem -Namespace $Namespace -EntryID $recs[$Index - 1].EntryID)
    }
    throw "Provide -EntryID (preferred) or -Index to choose a note."
}

# Record for one NoteItem (same fields as Get-NoteRecords, plus the full body with -WithBody).
function ConvertTo-NoteRecord($Note, [switch]$WithBody) {
    $body = [string]$Note.Body
    $rec = [ordered]@{
        EntryID    = $Note.EntryID
        Color      = $script:NoteColorNames[[int]$Note.Color]
        Title      = [string]$Note.Subject
        Preview    = $(if ($body.Length -gt 255) { $body.Substring(0, 255) } else { $body })
        Categories = [string]$Note.Categories
        Created    = $Note.CreationTime
        Modified   = $Note.LastModificationTime
    }
    if ($WithBody) { $rec.Body = $body }
    return $rec
}

# JSON copy of a note record: dates as ISO 8601.
function ConvertTo-NoteJson($Record) {
    $o = [ordered]@{}
    foreach ($k in $Record.Keys) { $o[$k] = $Record[$k] }
    $o.Created = Format-MailDateIso $Record.Created
    $o.Modified = Format-MailDateIso $Record.Modified
    return $o
}
