param(
    [Parameter(Mandatory=$true)]
    [string]$To,

    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [Parameter(Mandatory=$true)]
    [string]$Body,

    [Parameter(Mandatory=$true)]
    [string]$Account,

    [string]$CC = "",
    [string]$BCC = "",
    [string[]]$Attachment = @(),
    [switch]$HTML,
    [switch]$Confirm,
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Send As (From Specific Account) Script
# Usage: .\outlook-send-as.ps1 -Account "work@company.com" -To "email@example.com" -Subject "Hi" -Body "text" -Confirm
# Use outlook-accounts-list.ps1 to see available accounts

# Preview (shown always, no COM needed)
$format = if ($HTML) { "HTML" } else { "Plain Text" }
Write-Host "`n=== EMAIL TO SEND ===" -ForegroundColor Yellow
Write-Host "From (Account): $Account" -ForegroundColor Cyan
Write-Host "To: $To"
if ($CC) { Write-Host "CC: $CC" }
if ($BCC) { Write-Host "BCC: $BCC" }
Write-Host "Subject: $Subject"
Write-Host "Format: $format"
$bodyPreview = if ($Body.Length -gt 200) { $Body.Substring(0, 200) + "..." } else { $Body }
if ($StripLinks) { $bodyPreview = Strip-Links $bodyPreview }
Write-Host "Body preview: $bodyPreview"

# Validate attachments in preview
$validAttachments = @()
foreach ($file in $Attachment) {
    if (-not $file) { continue }
    if (Test-Path $file) {
        $validAttachments += (Resolve-Path $file).Path
        Write-Host "Attachment: $(Split-Path $file -Leaf)" -ForegroundColor Gray
    } else {
        Write-Host "Warning: Attachment not found - $file" -ForegroundColor Yellow
    }
}
Write-Host "========================" -ForegroundColor Yellow

if (-not $Confirm) {
    Write-Host "`n[SAFETY] Email NOT sent. Add -Confirm flag to actually send." -ForegroundColor Red
    Write-Host "Example: .\outlook-send-as.ps1 -Account '$Account' -To '$To' -Subject '$Subject' -Body '...' -Confirm" -ForegroundColor Gray
} else {
    $Outlook = $null
    $Mail = $null
    try {
        try {
            $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            $Outlook = New-Object -ComObject Outlook.Application
            Start-Sleep -Milliseconds 500
        }

        # Find the specified account
        $sendAccount = $null
        foreach ($acc in $Outlook.Session.Accounts) {
            if ($acc.SmtpAddress -eq $Account -or $acc.DisplayName -eq $Account) {
                $sendAccount = $acc
                break
            }
        }

        if (-not $sendAccount) {
            Write-Host "`nAccount '$Account' not found." -ForegroundColor Red
            Write-Host "`nAvailable accounts:" -ForegroundColor Yellow
            foreach ($acc in $Outlook.Session.Accounts) {
                Write-Host "  - $($acc.DisplayName) <$($acc.SmtpAddress)>" -ForegroundColor Gray
            }
            Write-Host "`nUse outlook-accounts-list.ps1 to see all accounts." -ForegroundColor Gray
        } else {
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

            $Mail.SendUsingAccount = $sendAccount

            foreach ($filePath in $validAttachments) {
                $Mail.Attachments.Add($filePath) | Out-Null
            }

            $Mail.Send()

            Write-Host "`n=== EMAIL SENT ===" -ForegroundColor Green
            Write-Host "From: $($sendAccount.DisplayName) <$($sendAccount.SmtpAddress)>" -ForegroundColor Cyan
            Write-Host "To: $To"
            if ($CC) { Write-Host "CC: $CC" }
            if ($BCC) { Write-Host "BCC: $BCC" }
            Write-Host "Subject: $Subject"
            Write-Host "Format: $format"
            if ($validAttachments.Count -gt 0) {
                $fileNames = $validAttachments | ForEach-Object { Split-Path $_ -Leaf }
                Write-Host "Attachments: $($fileNames -join ', ')" -ForegroundColor Gray
            }
        }
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
}
