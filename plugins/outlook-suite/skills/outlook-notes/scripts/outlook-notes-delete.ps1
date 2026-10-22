param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [ValidateSet("All", "Blue", "Green", "Pink", "Yellow", "White")]
    [string]$Color = "All",
    [string]$Search = "",
    [string]$Category = "",
    $StripLinks = $true,   # $true/$false (also "false" or 0 from pwsh -File)
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Note Delete: move a note to Deleted Items.
# Irreversible action: previews unless -Confirm (or confirmation for "delete" is turned off;
# see the outlook-mail SKILL.md "Confirmation").
# Usage: .\outlook-notes-delete.ps1 -EntryID "00000000..." -Confirm (preferred)
# Usage: .\outlook-notes-delete.ps1 -Index 1 -Confirm (position in outlook-notes-list.ps1 order, same filters)

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_notes.ps1")
$StripLinks = ConvertTo-StripLinks $StripLinks

try {
    $null = Connect-Outlook
    $note = Resolve-NoteItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Color $Color -Search $Search -Category $Category
    $rec = ConvertTo-NoteRecord $note
    $preview = Format-NotePreview $rec.Preview $StripLinks

    # Fail closed: anything but exactly $true (confirmed, or the gate turned off) means preview only.
    $go = [bool]$Confirm -or -not (Test-ConfirmRequired -Action "delete")
    if ($go -ne $true) {
        if ($Json) {
            $o = ConvertTo-NoteJson $rec
            $o.Preview = $preview
            $o.Action = "delete"; $o.Result = "preview"; $o.Done = $false
            $o.Hint = "Nothing was done. Run the same command with -Confirm to delete."
            Write-Json $o
            exit 0
        }
        Write-Host "=== DELETE NOTE (move to Deleted Items) (PREVIEW) ===" -ForegroundColor Yellow
        Write-Host "  EntryID: $($rec.EntryID)"
        Write-Host "  Color: $($rec.Color)"
        Write-Host "  Content: $preview" -ForegroundColor Gray
        Write-Host "  Modified: $(Format-MailDate $rec.Modified)" -ForegroundColor Gray
        $null = Approve-Action -Action "delete" -Confirmed $false -What "the delete"   # preview footer only; the gate decision is $go
        exit 0
    }

    $note.Delete()
    if ($Json) {
        $o = ConvertTo-NoteJson $rec
        $o.Preview = $preview
        $o.Action = "delete"; $o.Result = "deleted"
        Write-Json $o
        exit 0
    }
    Write-Host "=== NOTE DELETED (moved to Deleted Items) ===" -ForegroundColor Green
    Write-Host "  EntryID (no longer valid): $($rec.EntryID)"
    Write-Host "  Color: $($rec.Color)"
    Write-Host "  Content: $preview" -ForegroundColor Gray
}
catch {
    Stop-WithError $_
}
