param(
    [Parameter(Mandatory = $true)]
    [string]$To,
    [Parameter(Mandatory = $true)]
    [string]$Subject,
    [string]$Body = "",
    [string]$CC = "",
    [string]$BCC = "",
    [string[]]$Attachment = @(),
    [switch]$HTML,
    [string]$HtmlFile = "",
    [switch]$Native,
    [string]$Account = "",
    [ValidateSet("", "Low", "Normal", "High")]
    [string]$Importance = "",
    [ValidateSet("", "Normal", "Personal", "Private", "Confidential")]
    [string]$Sensitivity = "",
    [switch]$ReadReceipt,
    [switch]$DeliveryReceipt,
    [switch]$Confirm,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Send: compose and send an email.
# Outward-facing: previews unless -Confirm (or confirmation for "send" is turned off).
# Body options as in outlook-draft.ps1 (-HTML, -HtmlFile, -Native). Recipients are resolved
# first; nothing is sent if any name is unknown.
# Usage: .\outlook-send.ps1 -To "a@x.com" -Subject "Hello" -Body "Message" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    Invoke-NewMail -P $PSBoundParameters -Action "send" -Send
}
catch {
    Stop-WithError $_
}
