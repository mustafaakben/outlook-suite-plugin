param(
    [Parameter(Mandatory=$true)]
    [string]$To,

    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [Parameter(Mandatory=$true)]
    [string]$Body,

    [Parameter(Mandatory=$true)]
    [datetime]$SendAt,

    [string]$CC = "",
    [string]$BCC = "",
    [string[]]$Attachment = @(),
    [switch]$HTML,
    [switch]$ReadReceipt,
    [switch]$DeliveryReceipt,
    [switch]$Confirm
)

# Outlook Schedule Send Script
# Usage: .\outlook-schedule-send.ps1 -To "email@example.com" -Subject "Hello" -Body "Message" -SendAt "2026-02-10 09:00" -Confirm
# With receipts: .\outlook-schedule-send.ps1 -To "email@example.com" -Subject "Hi" -Body "Text" -SendAt "2026-02-10 09:00" -ReadReceipt -Confirm

# Validate SendAt is in the future
if ($SendAt -le (Get-Date)) {
    Write-Host "`nError: SendAt time must be in the future." -ForegroundColor Red
    Write-Host "Current time: $(Get-Date -Format 'g')" -ForegroundColor Gray
    Write-Host "You specified: $($SendAt.ToString('g'))" -ForegroundColor Gray
} else {
    # Show preview
    $formatType = if ($HTML) { "HTML" } else { "Plain Text" }
    $bodyPreview = if ($Body.Length -gt 200) { $Body.Substring(0, 200) + "..." } else { $Body }

    Write-Host "`n=== SCHEDULED EMAIL PREVIEW ===" -ForegroundColor Yellow
    Write-Host "To: $To"
    if ($CC) { Write-Host "CC: $CC" }
    if ($BCC) { Write-Host "BCC: $BCC" }
    Write-Host "Subject: $Subject"
    Write-Host "Format: $formatType" -ForegroundColor Gray
    Write-Host "Scheduled: $($SendAt.ToString('ddd, MMM d, yyyy h:mm tt'))" -ForegroundColor Cyan

    $timeUntil = $SendAt - (Get-Date)
    if ($timeUntil.TotalHours -ge 24) {
        Write-Host "  (in $([math]::Floor($timeUntil.TotalDays)) days, $($timeUntil.Hours) hours)" -ForegroundColor Gray
    } elseif ($timeUntil.TotalHours -ge 1) {
        Write-Host "  (in $([math]::Floor($timeUntil.TotalHours)) hours, $($timeUntil.Minutes) minutes)" -ForegroundColor Gray
    } else {
        Write-Host "  (in $($timeUntil.Minutes) minutes)" -ForegroundColor Gray
    }

    if ($ReadReceipt) { Write-Host "Read Receipt: REQUESTED" -ForegroundColor Cyan }
    if ($DeliveryReceipt) { Write-Host "Delivery Receipt: REQUESTED" -ForegroundColor Cyan }

    foreach ($att in $Attachment) {
        if (Test-Path $att) {
            Write-Host "Attachment: $att" -ForegroundColor Gray
        } else {
            Write-Host "Attachment: $att [NOT FOUND]" -ForegroundColor Red
        }
    }

    Write-Host "Body: $bodyPreview" -ForegroundColor Gray
    Write-Host "================================" -ForegroundColor Yellow

    if (-not $Confirm) {
        Write-Host "`n[SAFETY] Email NOT scheduled. Add -Confirm flag to actually schedule." -ForegroundColor Red
        Write-Host "Example: .\outlook-schedule-send.ps1 -To '$To' -Subject '$Subject' -Body '...' -SendAt '$($SendAt.ToString('g'))' -Confirm" -ForegroundColor Gray
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

            # Set deferred delivery time
            $Mail.DeferredDeliveryTime = $SendAt

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

            # Send (will be held in Outbox until SendAt time)
            $Mail.Send()

            Write-Host "`n=== EMAIL SCHEDULED ===" -ForegroundColor Green
            Write-Host "To: $To"
            Write-Host "Subject: $Subject"
            Write-Host "Scheduled for: $($SendAt.ToString('ddd, MMM d, yyyy h:mm tt'))" -ForegroundColor Cyan
            if ($ReadReceipt) { Write-Host "Read receipt: Requested" -ForegroundColor Gray }
            if ($DeliveryReceipt) { Write-Host "Delivery receipt: Requested" -ForegroundColor Gray }
            Write-Host "`nEmail is in Outbox and will send automatically." -ForegroundColor Yellow
            Write-Host "Keep Outlook running until the scheduled time." -ForegroundColor Yellow
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
