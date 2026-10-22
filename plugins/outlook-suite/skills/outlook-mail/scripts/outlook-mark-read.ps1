param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [switch]$Unread,
    [switch]$Json
)

# Outlook Mark Read: mark an email read (default) or unread (-Unread).
# Usage: .\outlook-mark-read.ps1 -EntryID "00000000..."
# Usage: .\outlook-mark-read.ps1 -EntryID "00000000..." -Unread
# Many emails at once: outlook-bulk.ps1 -Action MarkRead

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $want = [bool]$Unread
    $was = [bool]$email.UnRead
    if ($was -ne $want) { $email.UnRead = $want; $email.Save() }
    $state = if ($want) { "unread" } else { "read" }
    $title = if ($was -eq $want) { "EMAIL ALREADY $($state.ToUpper())" } else { "EMAIL MARKED $($state.ToUpper())" }
    Write-ActionResult -Title $title -Item $email -Changes ([ordered]@{ Status = $state }) -Json:$Json -Action "mark-read"
}
catch {
    Stop-WithError $_
}
