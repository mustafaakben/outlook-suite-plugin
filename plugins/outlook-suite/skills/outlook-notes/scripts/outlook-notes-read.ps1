param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [ValidateSet("All", "Blue", "Green", "Pink", "Yellow", "White")]
    [string]$Color = "All",
    [string]$Search = "",
    [string]$Category = "",
    $StripLinks = $true,   # $true/$false (also "false" or 0 from pwsh -File)
    [switch]$Json
)

# Outlook Note Read: the full text of one note.
# Links are shown as [URL-001] with a list of full URLs at the end; -StripLinks $false keeps them inline.
# Usage: .\outlook-notes-read.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-notes-read.ps1 -Index 1 (position in outlook-notes-list.ps1 order, same filters)
# With filters: .\outlook-notes-read.ps1 -Index 1 -Color Yellow -Search "budget"

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_notes.ps1")
$StripLinks = ConvertTo-StripLinks $StripLinks

try {
    $null = Connect-Outlook
    $note = Resolve-NoteItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Color $Color -Search $Search -Category $Category
    $rec = ConvertTo-NoteRecord $note -WithBody
    $links = Format-MailLinks -Text $rec.Body -Mode $(if ($StripLinks) { "domain" } else { "none" })
    if ($StripLinks) { $rec.Body = $links.Text }

    if ($Json) {
        $o = ConvertTo-NoteJson $rec
        $o.Remove("Preview")
        $o.Links = @($links.Links)
        Write-Json $o
        exit 0
    }

    $console = $script:NoteColorConsole[$script:NoteColorMap[$rec.Color]]
    Write-Host "=== NOTE ===" -ForegroundColor Cyan
    Write-Host "  EntryID: $($rec.EntryID)"
    Write-Host "  Color: $($rec.Color)" -ForegroundColor $console
    if ($rec.Categories) { Write-Host "  Categories: $($rec.Categories)" -ForegroundColor Green }
    Write-Host "  Created: $(Format-MailDate $rec.Created)" -ForegroundColor Gray
    Write-Host "  Modified: $(Format-MailDate $rec.Modified) ($(Format-Ago $rec.Modified))" -ForegroundColor Gray
    Write-Host "`n--- Content ---" -ForegroundColor Cyan
    Write-Host $(if ($rec.Body) { $rec.Body } else { "(Empty note)" }) -ForegroundColor $console
    if ($StripLinks -and $links.Links.Count -gt 0) {
        Write-Host "`n--- Links ---" -ForegroundColor Cyan
        foreach ($l in $links.Links) { Write-Host "  [$($l.Ref)] $($l.Url)" -ForegroundColor DarkGray }
    }
    $lineCount = if ($rec.Body) { ($rec.Body -split "`n").Count } else { 0 }
    Write-Host "`n($($rec.Body.Length) characters, $lineCount lines)" -ForegroundColor Gray
}
catch {
    Stop-WithError $_
}
