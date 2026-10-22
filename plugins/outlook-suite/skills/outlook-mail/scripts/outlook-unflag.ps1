param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [switch]$Complete,
    [switch]$Json
)

# Outlook Unflag: remove the follow-up flag, or mark it complete (-Complete).
# Usage: .\outlook-unflag.ps1 -EntryID "00000000..."
# Usage: .\outlook-unflag.ps1 -EntryID "00000000..." -Complete

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $before = switch ([int]$email.FlagStatus) { 2 { "Flagged" } 1 { "Complete" } default { "Not flagged" } }
    if ($Complete) {
        try { $email.MarkAsTask(5) }   # olMarkComplete
        catch {
            if (-not $email.Sent) { Stop-WithError "Outlook cannot mark a draft's flag complete (only sent or received mail). Clear it instead: run without -Complete." }
            throw
        }
        $email.Save()
        Write-ActionResult -Title "FLAG MARKED COMPLETE" -Item $email -Changes ([ordered]@{ Before = $before; Now = "Complete" }) -Json:$Json -Action "unflag"
    } else {
        if ($before -ne "Not flagged") {
            try { $email.ClearTaskFlag() } catch { }
            $email.Save()
            # ClearTaskFlag can silently do nothing (seen on drafts); verify and fall back.
            $email = $Namespace.GetItemFromID($email.EntryID)
            if ([int]$email.FlagStatus -ne 0) {
                $email.FlagStatus = 0
                $email.ReminderSet = $false
                $email.Save()
                $email = $Namespace.GetItemFromID($email.EntryID)
            }
            if ([int]$email.FlagStatus -ne 0) { Stop-WithError "Outlook did not clear the flag on this item." }
        }
        Write-ActionResult -Title $(if ($before -eq "Not flagged") { "EMAIL WAS NOT FLAGGED" } else { "FLAG REMOVED" }) -Item $email -Changes ([ordered]@{ Before = $before; Now = "Not flagged" }) -Json:$Json -Action "unflag"
    }
}
catch {
    Stop-WithError $_
}
