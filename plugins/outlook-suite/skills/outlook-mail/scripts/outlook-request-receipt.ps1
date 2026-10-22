param(
    [Parameter(Mandatory = $true)]
    [string]$To,
    [Parameter(Mandatory = $true)]
    [string]$Subject,
    [string]$Body = "",
    [switch]$ReadReceipt,
    [switch]$DeliveryReceipt,
    [string]$CC = "",
    [string]$BCC = "",
    [string[]]$Attachment = @(),
    [switch]$HTML,
    [string]$HtmlFile = "",
    [switch]$Native,
    [switch]$Confirm,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Request Receipt: send an email asking for a read and/or delivery receipt
# (default: read receipt). Outward-facing: previews unless -Confirm.
# Usage: .\outlook-request-receipt.ps1 -To "a@x.com" -Subject "Hi" -Body "text" -ReadReceipt -DeliveryReceipt -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $p = @{} + $PSBoundParameters
    if (-not $ReadReceipt -and -not $DeliveryReceipt) { $p.ReadReceipt = $true }
    Invoke-NewMail -P $p -Action "send" -Send
}
catch {
    Stop-WithError $_
}
