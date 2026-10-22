param(
    [int]$Depth = 3,
    [string]$Under = "",
    [switch]$AllStores,
    [switch]$NoCounts,
    [switch]$Json
)

# Outlook Folders: mail folders with paths you can pass to -Folder in other scripts.
# Default: your mailboxes (public folders skipped; add -AllStores), 3 levels deep.
# Usage: .\outlook-folders.ps1
# Usage: .\outlook-folders.ps1 -Under "Inbox" -Depth 5
# Usage: .\outlook-folders.ps1 -Json

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $rows = New-Object System.Collections.Generic.List[object]

    function Add-FolderRows($folder, [string]$path, [int]$level, [string]$storeName) {
        foreach ($f in $folder.Folders) {
            $p = if ($path) { "$path/$($f.Name)" } else { $f.Name }
            $row = [ordered]@{ Path = $p; Mailbox = $storeName; Level = $level; EntryID = $f.EntryID; StoreID = $f.StoreID; Items = $null; Unread = $null; Subfolders = $f.Folders.Count }
            if (-not $NoCounts) {
                try { $row.Items = $f.Items.Count } catch { }
                try { $row.Unread = $f.UnReadItemCount } catch { }
            }
            $rows.Add($row)
            if ($level + 1 -lt $Depth -and $row.Subfolders -gt 0) { Add-FolderRows $f $p ($level + 1) $storeName }
        }
    }

    if ($Under) {
        $start = Resolve-OutlookFolder -Namespace $Namespace -Path $Under -AllStores:$AllStores
        $base = ($start.FolderPath -replace '^\\\\[^\\]+\\', '') -replace '\\', '/'
        Add-FolderRows $start $base 0 $start.Store.DisplayName
    } else {
        foreach ($s in (Get-SearchStores -Namespace $Namespace -AllStores:$AllStores)) {
            Add-FolderRows $s.GetRootFolder() "" 0 $s.DisplayName
        }
    }

    if ($Json) {
        Write-Json ([ordered]@{ count = $rows.Count; folders = $rows })
        exit 0
    }

    Write-Host "=== OUTLOOK FOLDERS ===" -ForegroundColor Cyan
    Write-Host "Use the path with -Folder, e.g. -Folder ""Inbox/Projects"". For another mailbox: -Folder ""<Mailbox>/Inbox""." -ForegroundColor DarkGray
    $mailbox = ""
    foreach ($r in $rows) {
        if ($r.Mailbox -ne $mailbox) {
            $mailbox = $r.Mailbox
            Write-Host ""
            Write-Host "[$mailbox]" -ForegroundColor Green
        }
        $indent = "  " * ($r.Level + 1)
        $counts = if ($NoCounts) { "" } else { "  ($($r.Items) items, $($r.Unread) unread)" }
        $color = if ($r.Unread -gt 0) { "Yellow" } else { "Gray" }
        Write-Host "$indent$($r.Path)$counts" -ForegroundColor $color
    }
    Write-Host ""
    Write-Host "=== $($rows.Count) folders (depth $Depth) ===" -ForegroundColor Cyan
    Write-Host "Folder EntryIDs: add -Json." -ForegroundColor DarkGray
}
catch {
    Stop-WithError $_
}
