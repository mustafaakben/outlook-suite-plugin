param(
    [ValidateRange(1, 2147483647)]
    [int]$Limit = 20,
    [ValidateSet("All", "Blue", "Green", "Pink", "Yellow", "White")]
    [string]$Color = "All",
    [string]$Search = "",
    [string]$Category = "",
    $StripLinks = $true,   # $true/$false (also "false" or 0 from pwsh -File)
    [switch]$Json
)

# Outlook Notes List: sticky notes, most recently changed first (one fast table query).
# Usage: .\outlook-notes-list.ps1
# With limit: .\outlook-notes-list.ps1 -Limit 50
# Filter by color: .\outlook-notes-list.ps1 -Color Yellow
# Search the whole note text (substring, any case): .\outlook-notes-list.ps1 -Search "meeting"
# Category: .\outlook-notes-list.ps1 -Category "Research"
# JSON: .\outlook-notes-list.ps1 -Search "meeting" -Json   (Preview = first 255 characters)

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_notes.ps1")
$StripLinks = ConvertTo-StripLinks $StripLinks

try {
    $null = Connect-Outlook
    $recs = Get-NoteRecords -Namespace $Namespace -Color $Color -Search $Search -Category $Category
    $shown = @($recs | Select-Object -First $Limit)

    if ($Json) {
        $items = @(foreach ($r in $shown) {
            $o = ConvertTo-NoteJson $r
            if ($StripLinks) { $o.Preview = (Format-MailLinks -Text $o.Preview -Mode domain).Text }
            $o
        })
        Write-Json ([ordered]@{ folder = $Namespace.GetDefaultFolder(12).FolderPath; total = $recs.Count; count = $items.Count
            query = [ordered]@{ Color = $Color; Search = $Search; Category = $Category }; items = $items })
        exit 0
    }

    Write-Host "=== OUTLOOK NOTES ===" -ForegroundColor Cyan
    $criteria = @()
    if ($Color -ne "All") { $criteria += "color: $Color" }
    if ($Search) { $criteria += "text contains '$Search'" }
    if ($Category) { $criteria += "category: $Category" }
    if ($criteria.Count) { Write-Host "Filters: $($criteria -join ' | ')" -ForegroundColor Gray }
    Write-Host ""

    $n = 0
    foreach ($r in $shown) {
        $n++
        $console = $script:NoteColorConsole[$script:NoteColorMap[$r.Color]]
        Write-Host "$n. " -NoNewline
        Write-Host "[$($r.Color)] " -NoNewline -ForegroundColor $console
        Write-Host (Format-NotePreview $r.Preview $StripLinks) -ForegroundColor $console
        Write-Host "      EntryID: $($r.EntryID)" -ForegroundColor DarkGray
        if ($r.Categories) { Write-Host "      Categories: $($r.Categories)" -ForegroundColor Green }
        Write-Host "      Modified: $(Format-MailDate $r.Modified) ($(Format-Ago $r.Modified))" -ForegroundColor Gray
        Write-Host ""
    }
    if ($n -eq 0) {
        Write-Host $(if ($criteria.Count) { "No notes found matching criteria." } else { "No notes found." }) -ForegroundColor Gray
    } else {
        $more = if ($recs.Count -gt $n) { " of $($recs.Count) matching; raise -Limit for more" } else { "" }
        Write-Host "=== Showing $n$more ===" -ForegroundColor Cyan
        Write-Host "Next: outlook-notes-read.ps1 -EntryID <EntryID>" -ForegroundColor DarkGray
    }
}
catch {
    Stop-WithError $_
}
