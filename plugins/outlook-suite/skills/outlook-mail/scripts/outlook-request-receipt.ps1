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
    [switch]$ReadReceipt,
    [switch]$DeliveryReceipt,
    [switch]$Confirm,
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Send with Receipt Request Script
# Usage: .\outlook-request-receipt.ps1 -To "email@example.com" -Subject "Hi" -Body "text" -ReadReceipt -Confirm
# Both receipts: .\outlook-request-receipt.ps1 -To "email@example.com" -Subject "Hi" -Body "text" -ReadReceipt -DeliveryReceipt -Confirm

if (-not $ReadReceipt -and -not $DeliveryReceipt) {
    Write-Host "`nPlease specify at least one receipt type:" -ForegroundColor Red
    Write-Host "  -ReadReceipt      (notification when email is read)" -ForegroundColor Gray
    Write-Host "  -DeliveryReceipt  (notification when email is delivered)" -ForegroundColor Gray
} else {
    $formatType = if ($HTML) { "HTML" } else { "Plain Text" }
    $bodyPreview = if ($Body.Length -gt 200) { $Body.Substring(0, 200) + "..." } else { $Body }

    Write-Host "`n=== EMAIL TO SEND ===" -ForegroundColor Yellow
    Write-Host "To: $To"
    if ($CC) { Write-Host "CC: $CC" }
    if ($BCC) { Write-Host "BCC: $BCC" }
    Write-Host "Subject: $Subject"
    Write-Host "Format: $formatType" -ForegroundColor Gray
    if ($ReadReceipt) { Write-Host "Read Receipt: REQUESTED" -ForegroundColor Cyan }
    if ($DeliveryReceipt) { Write-Host "Delivery Receipt: REQUESTED" -ForegroundColor Cyan }

    foreach ($att in $Attachment) {
        if (Test-Path $att) {
            Write-Host "Attachment: $att" -ForegroundColor Gray
        } else {
            Write-Host "Attachment: $att [NOT FOUND]" -ForegroundColor Red
        }
    }

    if ($StripLinks) { $bodyPreview = Strip-Links $bodyPreview }
    Write-Host "Body: $bodyPreview" -ForegroundColor Gray
    Write-Host "========================" -ForegroundColor Yellow

    if (-not $Confirm) {
        Write-Host "`n[SAFETY] Email NOT sent. Add -Confirm flag to actually send." -ForegroundColor Red
        Write-Host "Example: .\outlook-request-receipt.ps1 -To '$To' -Subject '$Subject' -Body '...' -ReadReceipt -Confirm" -ForegroundColor Gray
    } else {
        $Outlook = $null
        $Mail = $null

        try {
            # Connect to Outlook - try active instance first
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

            # Set receipt options
            if ($ReadReceipt) { $Mail.ReadReceiptRequested = $true }
            if ($DeliveryReceipt) { $Mail.OriginatorDeliveryReportRequested = $true }

            # Add attachments
            foreach ($att in $Attachment) {
                if (Test-Path $att) {
                    $resolvedPath = (Resolve-Path $att).Path
                    $Mail.Attachments.Add($resolvedPath) | Out-Null
                }
            }

            $Mail.Send()

            Write-Host "`n=== EMAIL SENT ===" -ForegroundColor Green
            Write-Host "To: $To"
            Write-Host "Subject: $Subject"
            if ($ReadReceipt) { Write-Host "Read receipt will be sent when recipient opens the email." -ForegroundColor Gray }
            if ($DeliveryReceipt) { Write-Host "Delivery receipt will be sent when email reaches recipient's server." -ForegroundColor Gray }
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
}
