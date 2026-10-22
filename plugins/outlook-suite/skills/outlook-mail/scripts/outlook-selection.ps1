param(
    [switch]$Body,
    [ValidateSet("domain", "full", "none")]
    [string]$Links = "domain",
    [int]$Limit = 50,
    [switch]$Json
)

# Outlook Selection: the email(s) currently selected in the Outlook window, so you can say
# "this email" and the agent knows which one. Prints EntryIDs for other scripts.
# -Body adds each email's full text (links as [URL-001] references).
# Usage: .\outlook-selection.ps1
# Usage: .\outlook-selection.ps1 -Body

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $exp = $null; try { $exp = $Outlook.ActiveExplorer() } catch { }
    if (-not $exp) { Stop-WithError "No Outlook window is open (Outlook may be running in the background only)." }
    $sel = $exp.Selection
    $folder = ""; try { $folder = $exp.CurrentFolder.FolderPath } catch { }
    $records = New-Object System.Collections.Generic.List[object]
    $other = 0
    for ($i = 1; $i -le $sel.Count -and $records.Count -lt $Limit; $i++) {
        $it = $sel.Item($i)
        if ([int]$it.Class -ne 43) { $other++; continue }   # 43 = olMail
        $rec = if ($Body) { ConvertTo-MailRecord $it -WithRecipients -WithBody -Links $Links } else { ConvertTo-MailRecord $it -WithRecipients }
        $records.Add($rec)
    }
    if ($Json) { Write-Json ([ordered]@{ folder = $folder; selected = $sel.Count; emails = @($records | ForEach-Object { ConvertTo-JsonRecord $_ }); otherItems = $other }); exit 0 }
    Write-Host "=== SELECTED IN OUTLOOK ===" -ForegroundColor Cyan
    Write-Host "Folder: $folder" -ForegroundColor Gray
    if ($sel.Count -eq 0) { Write-Host "Nothing is selected (a group header or an empty folder may be selected)." -ForegroundColor Gray; exit 0 }
    $n = 0
    foreach ($r in $records) { $n++; Write-Host ""; Write-MailRecord $r -Header "--- Selected $n of $($records.Count) ---" -LinkTable:$Body }
    if ($other) { Write-Host "($other selected item(s) are not emails: meeting requests, reports, etc.)" -ForegroundColor Gray }
}
catch {
    Stop-WithError $_
}
