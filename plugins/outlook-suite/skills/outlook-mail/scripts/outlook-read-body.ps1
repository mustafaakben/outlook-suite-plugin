param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [ValidateSet("domain", "full", "none")]
    [string]$Links = "domain",
    [ValidateSet("text", "html")]
    [string]$Format = "text",
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy: -StripLinks:$false is the same as -Links full
)

# Outlook Email Body Reader: one email in full (headers, recipients, body, attachments, links).
# Links appear as [URL-001] with a reference table; full URLs: -Links full or outlook-links.ps1.
# Usage: .\outlook-read-body.ps1 -EntryID "00000000..."
# Usage: .\outlook-read-body.ps1 -EntryID "00000000..." -Links full
# Usage: .\outlook-read-body.ps1 -EntryID "00000000..." -Format html      (raw HTML body)

. (Join-Path $PSScriptRoot "_common.ps1")
if (-not $StripLinks) { $Links = "full" }

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days

    if ($Format -eq "html") {
        $rec = ConvertTo-MailRecord $email -WithRecipients
        $rec.HTMLBody = $email.HTMLBody
        if ($Json) { Write-Json (ConvertTo-JsonRecord $rec); exit 0 }
        Write-MailRecord $rec -Header "=== EMAIL (HTML) ==="
        Write-Host "--- HTML BODY ---" -ForegroundColor Gray
        Write-Host $rec.HTMLBody
        Write-Host "--- END ---" -ForegroundColor Gray
        exit 0
    }

    $rec = ConvertTo-MailRecord $email -WithRecipients -WithBody -Links $Links
    if ($Json) { Write-Json (ConvertTo-JsonRecord $rec); exit 0 }
    Write-MailRecord $rec -Header "=== EMAIL ===" -LinkTable
    Write-Host "--- END ---" -ForegroundColor Gray
}
catch {
    Stop-WithError $_
}
