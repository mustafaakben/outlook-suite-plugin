param(
    [Parameter(Mandatory=$true)]
    [string]$To,

    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [Parameter(Mandatory=$true)]
    [string]$Body,

    [string]$Options = "Yes;No",
    [string]$CC = "",
    [string]$BCC = "",
    [switch]$HTML,
    [switch]$Confirm,
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Voting Buttons Script
# Usage: .\outlook-voting.ps1 -To "team@example.com" -Subject "Lunch poll" -Body "Where should we eat?" -Options "Pizza;Tacos;Sushi" -Confirm
# Default options: Yes;No
# Custom options: .\outlook-voting.ps1 -To "x@y.com" -Subject "Meeting time?" -Body "Pick one" -Options "9am;10am;11am;None work" -Confirm
# Without -Confirm, only a preview is shown and nothing is sent.

$Outlook = $null
$Mail = $null

try {
    # Body preview for display
    $bodyPreview = $Body
    if ($bodyPreview.Length -gt 200) {
        $bodyPreview = $bodyPreview.Substring(0, 200) + "..."
    }

    # Format indicator
    $formatLabel = if ($HTML) { "HTML" } else { "Plain Text" }

    # Parse voting options
    $optionList = $Options -split ";"

    # Show preview
    Write-Host "`n=== VOTING EMAIL PREVIEW ===" -ForegroundColor Yellow
    Write-Host "To: $To"
    if ($CC) { Write-Host "CC: $CC" }
    if ($BCC) { Write-Host "BCC: $BCC" }
    Write-Host "Subject: $Subject"
    Write-Host "Format: $formatLabel" -ForegroundColor Gray
    Write-Host "`nVoting options:" -ForegroundColor Cyan
    $optionNum = 0
    foreach ($opt in $optionList) {
        $optionNum++
        Write-Host "  [$optionNum] $opt"
    }
    if ($StripLinks) { $bodyPreview = Strip-Links $bodyPreview }
    Write-Host "`nBody:" -ForegroundColor Gray
    Write-Host $bodyPreview
    Write-Host "==============================" -ForegroundColor Yellow

    if (-not $Confirm) {
        Write-Host "`n[SAFETY] Email NOT sent. Add -Confirm flag to actually send." -ForegroundColor Red
        Write-Host "Example: .\outlook-voting.ps1 -To '$To' -Subject '$Subject' -Body '...' -Options '$Options' -Confirm" -ForegroundColor Gray
    } else {
        # Connect to Outlook (fast path if already running)
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

        # Set voting options
        $Mail.VotingOptions = $Options

        $Mail.Send()

        Write-Host "`n=== VOTING EMAIL SENT ===" -ForegroundColor Green
        Write-Host "Recipients will see voting buttons in their email." -ForegroundColor Gray
        Write-Host "Responses will appear in your Inbox with their votes." -ForegroundColor Gray
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
