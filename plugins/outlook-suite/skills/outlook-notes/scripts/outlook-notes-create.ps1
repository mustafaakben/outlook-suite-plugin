param(
    [Parameter(Mandatory=$true)]
    [string]$Body,

    [ValidateSet("Blue", "Green", "Pink", "Yellow", "White")]
    [string]$Color = "Yellow",
    [string]$Categories = "",
    $StripLinks = $true,   # $true/$false (also "false" or 0 from pwsh -File)
    [switch]$Json
)

# Outlook Note Create: save a new sticky note (local item, nothing is sent).
# The first line of the text becomes the note's title.
# Usage: .\outlook-notes-create.ps1 -Body "Remember to call John"
# With color: .\outlook-notes-create.ps1 -Body "Important meeting notes" -Color Pink
# Category: .\outlook-notes-create.ps1 -Body "Idea: ..." -Categories "Research"

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_notes.ps1")
$StripLinks = ConvertTo-StripLinks $StripLinks

try {
    $null = Connect-Outlook
    $note = $Outlook.CreateItem(5)   # olNoteItem
    $note.Body = $Body
    if ($Categories) { $note.Categories = $Categories }   # before the color: setting categories resets it
    $note.Color = $script:NoteColorMap[$Color]
    # NoteItem.Color alone is not saved on a new note; write the color property itself too.
    $note.PropertyAccessor.SetProperty($script:NoteProp.Color, $script:NoteColorMap[$Color])
    $note.Save()

    $rec = ConvertTo-NoteRecord $note
    if ($Json) {
        $o = ConvertTo-NoteJson $rec
        $o.Action = "create"; $o.Result = "NOTE CREATED"
        if ($StripLinks) { $o.Preview = (Format-MailLinks -Text $o.Preview -Mode domain).Text }
        Write-Json $o
        exit 0
    }
    Write-Host "=== NOTE CREATED ===" -ForegroundColor Green
    Write-Host "  EntryID: $($rec.EntryID)"
    Write-Host "  Color: $Color" -ForegroundColor $script:NoteColorConsole[$script:NoteColorMap[$Color]]
    if ($rec.Categories) { Write-Host "  Categories: $($rec.Categories)" -ForegroundColor Green }
    Write-Host "  Content: $(Format-NotePreview $Body $StripLinks)" -ForegroundColor Gray
}
catch {
    Stop-WithError $_
}
