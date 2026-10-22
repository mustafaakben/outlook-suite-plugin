param(
    [Parameter(Mandatory = $true)]
    [string]$To,
    [Parameter(Mandatory = $true)]
    [string]$Subject,
    [string]$Body = "",
    [Parameter(Mandatory = $true)]
    [datetime]$SendAt,
    [string]$CC = "",
    [string]$BCC = "",
    [string[]]$Attachment = @(),
    [switch]$HTML,
    [string]$HtmlFile = "",
    [switch]$Native,
    [string]$Account = "",
    [switch]$ReadReceipt,
    [switch]$DeliveryReceipt,
    [switch]$Confirm,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Schedule Send: send at a later time (Outlook's Delay Delivery). The message waits
# in the Outbox until -SendAt (Exchange: on the server; other accounts: Outlook must be running).
# To cancel, delete it from the Outbox: outlook-find.ps1 -Folder Outbox, then outlook-delete.ps1.
# Outward-facing: previews unless -Confirm.
# Usage: .\outlook-schedule-send.ps1 -To "a@x.com" -Subject "Hello" -Body "Message" -SendAt "2026-11-02 09:00" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    if ($SendAt -le (Get-Date)) { Stop-WithError "-SendAt must be in the future (got $(Format-MailDate $SendAt))." }
    $null = Connect-Outlook
    Invoke-NewMail -P $PSBoundParameters -Action "send" -Send
}
catch {
    Stop-WithError $_
}
