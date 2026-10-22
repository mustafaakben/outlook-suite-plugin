param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [Parameter(Mandatory = $true)]
    [string]$Folder,
    [switch]$AllStores,
    [int]$Days = 7,
    [switch]$Json
)

# Outlook Move: move an email to a folder (name or path, e.g. "Archive", "Inbox/Projects",
# "someone@x.edu/Inbox"). Moving can change the EntryID: the new one is printed.
# Usage: .\outlook-move.ps1 -EntryID "00000000..." -Folder "Archive"
# Many emails at once: outlook-bulk.ps1 -Action Move

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder -AllStores:$AllStores
    $from = $email.Parent.FolderPath
    $oldId = $email.EntryID
    $moved = $email.Move($target)
    $changes = [ordered]@{ From = $from; To = $target.FolderPath }
    if ($moved.EntryID -ne $oldId) { $changes.OldEntryID = $oldId }
    Write-ActionResult -Title "EMAIL MOVED" -Item $moved -Changes $changes -Json:$Json -Action "move"
}
catch {
    Stop-WithError $_
}
