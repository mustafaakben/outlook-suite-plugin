param(
    [Parameter(Mandatory = $true)]
    [string]$EntryID,
    [string]$StoreID = "",
    [Parameter(Mandatory = $true)]
    [string]$Folder,
    [switch]$Json
)

# Outlook Copy: copy an email into a folder (the original stays). Prints the copy's EntryID.
# Usage: .\outlook-copy.ps1 -EntryID "00000000..." -Folder "Archive/2026"

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
    $copy = $email.Copy()   # lands in the email's own folder
    # Outlook cannot move an item into the folder it is already in: a copy into the same folder is done.
    $moved = if ($email.Parent.EntryID -eq $target.EntryID) { $copy } else { $copy.Move($target) }
    Write-ActionResult -Title "EMAIL COPIED" -Item $moved -Changes ([ordered]@{ Original = $email.EntryID; From = $email.Parent.FolderPath }) -Json:$Json -Action "copy"
}
catch {
    Stop-WithError $_
}
