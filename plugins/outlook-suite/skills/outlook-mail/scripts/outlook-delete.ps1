param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [switch]$Permanent,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Delete: move an email to Deleted Items, or purge it (-Permanent, cannot be undone).
# Irreversible action: previews unless -Confirm (or confirmation for "delete" is turned off;
# see SKILL.md "Confirmation").
# Usage: .\outlook-delete.ps1 -EntryID "00000000..." -Confirm
# Usage: .\outlook-delete.ps1 -EntryID "00000000..." -Permanent -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $mode = if ($Permanent) { "PERMANENT DELETE (cannot be undone)" } else { "DELETE (move to Deleted Items)" }
    if (-not (Confirm-ItemAction -Action "delete" -Title $mode -Item $email -Confirmed $Confirm -What "the delete")) { exit 0 }

    $rec = ConvertTo-MailRecord $email
    if (-not $Permanent) {
        $email.Delete()
        if ($Json) { Write-Json ([ordered]@{ Action = "delete"; Result = "deleted"; EntryID = $rec.EntryID; Subject = $rec.Subject; Folder = $rec.Folder; Permanent = $false }); exit 0 }
        Write-Host "=== EMAIL DELETED (moved to Deleted Items) ===" -ForegroundColor Green
    } else {
        # Purge = delete from Deleted Items. Move there first, then re-open by the new
        # EntryID so the exact message (and nothing else) is removed.
        $deleted = $Namespace.GetDefaultFolder(3)
        $inDeleted = $false
        try { $inDeleted = ($email.Parent.EntryID -eq $deleted.EntryID -and $email.Parent.StoreID -eq $deleted.StoreID) } catch { }
        $victim = if ($inDeleted) { $email } else { $email.Move($deleted) }
        if (-not $victim) { Stop-WithError "Could not move the message to Deleted Items; nothing was purged." }
        $victim = $Namespace.GetItemFromID($victim.EntryID, $deleted.StoreID)
        $victim.Delete()
        if ($Json) { Write-Json ([ordered]@{ Action = "delete"; Result = "purged"; EntryID = $rec.EntryID; Subject = $rec.Subject; Folder = $rec.Folder; Permanent = $true }); exit 0 }
        Write-Host "=== EMAIL PERMANENTLY DELETED ===" -ForegroundColor Green
    }
    Write-Host "  EntryID (no longer valid): $($rec.EntryID)"
    Write-Host "  Subject: $($rec.Subject)" -ForegroundColor Yellow
    Write-Host "  From: $($rec.From)"
    Write-Host "  Date: $(Format-MailDate $rec.Date)"
    Write-Host "  Was in: $($rec.Folder)" -ForegroundColor DarkGray
}
catch {
    Stop-WithError $_
}
