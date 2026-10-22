param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Create", "Rename", "Move", "Copy", "Delete", "Empty", "MarkAllRead")]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    [string]$Folder,
    [string]$Name = "",
    [string]$To = "",
    [switch]$AllStores,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Folder Manage: create, rename, move, copy, delete or empty a mail folder, or mark
# everything in it read. Paths as in outlook-folders.ps1 ("Inbox/Projects", "Mailbox/Inbox/X").
#   -Action Create -Folder "Inbox" -Name "Projects"          new subfolder "Inbox/Projects"
#   -Action Rename -Folder "Inbox/Projects" -Name "Work"
#   -Action Move   -Folder "Inbox/Projects" -To "Archive"     becomes "Archive/Projects"
#   -Action Copy   -Folder "Inbox/Projects" -To "Archive"     copy with its emails
#   -Action Delete -Folder "Inbox/Projects"                   moves it to Deleted Items
#   -Action Empty  -Folder "Inbox/Projects"                   moves all its emails to Deleted Items
#   -Action MarkAllRead -Folder "Inbox/Projects"
# Rename, Move, Delete, Empty and MarkAllRead are gated: preview unless -Confirm.
# Exchange does not let scripts permanently delete folders that are already in Deleted Items.

. (Join-Path $PSScriptRoot "_common.ps1")

function Get-FolderPathShort($f) { return ($f.FolderPath -replace '^\\\\', '') -replace '\\', '/' }

try {
    $null = Connect-Outlook
    $f = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder -AllStores:$AllStores
    $path = Get-FolderPathShort $f
    $gated = $Action -in "Rename", "Move", "Delete", "Empty", "MarkAllRead"
    $detail = [ordered]@{ Folder = $path; Items = $f.Items.Count; Unread = $f.UnReadItemCount; Subfolders = $f.Folders.Count }

    switch ($Action) {
        "Create" { if (-not $Name) { Stop-WithError "Give the new folder's -Name." } }
        "Rename" { if (-not $Name) { Stop-WithError "Give the new -Name." }; $detail.NewName = $Name }
        { $_ -in "Move", "Copy" } {
            if (-not $To) { Stop-WithError "Give the destination parent folder with -To." }
            $dest = Resolve-OutlookFolder -Namespace $Namespace -Path $To -AllStores:$AllStores
            $detail.Destination = Get-FolderPathShort $dest
        }
    }

    if ($gated -and -not ($Confirm -or -not (Test-ConfirmRequired -Action "folders"))) {
        Write-Host "=== FOLDER $($Action.ToUpper()) (PREVIEW) ===" -ForegroundColor Yellow
        foreach ($k in $detail.Keys) { Write-Host "  ${k}: $($detail[$k])" }
        Approve-Action -Action "folders" -Confirmed $false -What "this folder change" | Out-Null
        exit 0
    }

    $result = [ordered]@{ Action = $Action }
    switch ($Action) {
        "Create" {
            foreach ($c in $f.Folders) { if ($c.Name -ieq $Name) { Stop-WithError "'$path/$Name' already exists." } }
            $new = $f.Folders.Add($Name)
            $result.Folder = Get-FolderPathShort $new; $result.EntryID = $new.EntryID
        }
        "Rename" {
            foreach ($c in $f.Parent.Folders) { if ($c.Name -ieq $Name -and $c.EntryID -ne $f.EntryID) { Stop-WithError "A sibling folder named '$Name' already exists." } }
            $f.Name = $Name
            $result.Was = $path; $result.Folder = Get-FolderPathShort $f
        }
        "Move" {
            $f.MoveTo($dest)
            $moved = $null; foreach ($c in $dest.Folders) { if ($c.Name -eq $f.Name) { $moved = $c } }
            $result.Was = $path; $result.Folder = $(if ($moved) { Get-FolderPathShort $moved } else { "$($detail.Destination)/$($f.Name)" })
        }
        "Copy" {
            $copy = $f.CopyTo($dest)
            $result.From = $path; $result.Folder = Get-FolderPathShort $copy; $result.EntryID = $copy.EntryID
        }
        "Delete" {
            $f.Delete()
            $result.Deleted = $path; $result.Note = "Moved to Deleted Items with its emails and subfolders."
        }
        "Empty" {
            $n = 0
            for ($i = $f.Items.Count; $i -ge 1; $i--) { try { $f.Items.Item($i).Delete(); $n++ } catch { } }
            $result.Folder = $path; $result.Deleted = $n; $result.Left = $f.Items.Count
        }
        "MarkAllRead" {
            $n = 0
            $unread = $f.Items.Restrict("[UnRead] = True")
            for ($i = $unread.Count; $i -ge 1; $i--) { try { $it = $unread.Item($i); $it.UnRead = $false; $it.Save(); $n++ } catch { } }
            $result.Folder = $path; $result.MarkedRead = $n; $result.UnreadLeft = $f.UnReadItemCount
        }
    }

    if ($Json) { Write-Json $result; exit 0 }
    Write-Host "=== FOLDER $($Action.ToUpper()) DONE ===" -ForegroundColor Green
    foreach ($k in $result.Keys) { if ($k -ne "Action") { Write-Host "  ${k}: $($result[$k])" } }
}
catch {
    Stop-WithError $_
}
