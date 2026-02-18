param(
    [Parameter(Mandatory=$true)]
    [string]$To,

    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [Parameter(Mandatory=$true)]
    [string]$Body,

    [string]$CC = "",
    [string]$BCC = "",
    [string[]]$Attachment = @(),
    [switch]$HTML,
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Draft Creator
# Usage: .\outlook-draft.ps1 -To "email@example.com" -Subject "Hello" -Body "Message here"
# With attachment: .\outlook-draft.ps1 -To "email" -Subject "subj" -Body "text" -Attachment "C:\file.pdf"
# Multiple attachments: .\outlook-draft.ps1 -To "email" -Subject "subj" -Body "text" -Attachment "C:\a.pdf","C:\b.pdf"

$Outlook = $null
$Mail = $null
try {
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Mail = $Outlook.CreateItem(0)

    $Mail.To = $To
    $Mail.Subject = $Subject

    if ($HTML) {
        $Mail.HTMLBody = $Body
    } else {
        $Mail.Body = $Body
    }

    if ($CC) { $Mail.CC = $CC }
    if ($BCC) { $Mail.BCC = $BCC }

    # Attach files with validation
    $attachedFiles = @()
    foreach ($file in $Attachment) {
        if (-not $file) { continue }
        if (Test-Path $file) {
            $Mail.Attachments.Add((Resolve-Path $file).Path) | Out-Null
            $attachedFiles += (Split-Path $file -Leaf)
        } else {
            Write-Host "Warning: Attachment not found - $file" -ForegroundColor Yellow
        }
    }

    $Mail.Save()

    Write-Host "`n=== DRAFT CREATED ===" -ForegroundColor Green
    Write-Host "To: $To"
    if ($CC) { Write-Host "CC: $CC" }
    if ($BCC) { Write-Host "BCC: $BCC" }
    Write-Host "Subject: $Subject"
    $format = if ($HTML) { "HTML" } else { "Plain Text" }
    Write-Host "Format: $format"
    $bodyPreview = if ($Body.Length -gt 100) { $Body.Substring(0, 100) + "..." } else { $Body }
    if ($StripLinks) { $bodyPreview = Strip-Links $bodyPreview }
    Write-Host "Body preview: $bodyPreview"
    if ($attachedFiles.Count -gt 0) {
        Write-Host "Attachments: $($attachedFiles -join ', ')" -ForegroundColor Gray
    }
    Write-Host "`nDraft saved to Drafts folder." -ForegroundColor Cyan
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Mail) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
