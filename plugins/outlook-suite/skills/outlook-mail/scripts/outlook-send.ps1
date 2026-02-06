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
    [switch]$Confirm
)

# Outlook Email Sender
# Usage: .\outlook-send.ps1 -To "email@example.com" -Subject "Hello" -Body "Message" -Confirm
# IMPORTANT: Use -Confirm flag to actually send (safety measure)

# Preview (shown always, no COM needed)
$format = if ($HTML) { "HTML" } else { "Plain Text" }
Write-Host "`n=== EMAIL TO SEND ===" -ForegroundColor Yellow
Write-Host "To: $To"
if ($CC) { Write-Host "CC: $CC" }
if ($BCC) { Write-Host "BCC: $BCC" }
Write-Host "Subject: $Subject"
Write-Host "Format: $format"
if ($Body.Length -gt 200) {
    Write-Host "Body preview: $($Body.Substring(0, 200))..."
} else {
    Write-Host "Body preview: $Body"
}

# Show attachment list in preview
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
    Write-Host "Example: .\outlook-send.ps1 -To '$To' -Subject '$Subject' -Body '...' -Confirm" -ForegroundColor Gray
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

        foreach ($filePath in $validAttachments) {
            $Mail.Attachments.Add($filePath) | Out-Null
        }

        $Mail.Send()

        Write-Host "`n=== EMAIL SENT ===" -ForegroundColor Green
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
