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
    [switch]$Display,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Draft: create a new email in Drafts (nothing is sent). Prints the draft's EntryID.
# Body: plain text by default; -HTML (Body is HTML), -HtmlFile <path>, or -Native (plain text
# rendered in Outlook's own HTML style: Aptos 11pt div blocks).
# Recipients are resolved first; an unknown name fails before anything is saved.
# Usage: .\outlook-draft.ps1 -To "a@x.com" -Subject "Hello" -Body "Message"
# Usage: .\outlook-draft.ps1 -To "a@x.com" -Subject "Hi" -HtmlFile ".\body.html" -Attachment ".\report.pdf"
# Send it later: .\outlook-send-draft.ps1 -EntryID "<EntryID>" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    Invoke-NewMail -P $PSBoundParameters -Action "draft"
}
catch {
    Stop-WithError $_
}
