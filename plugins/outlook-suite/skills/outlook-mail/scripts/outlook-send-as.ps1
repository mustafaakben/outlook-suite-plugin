param(
    [Parameter(Mandatory = $true)]
    [string]$Account,
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
    [switch]$Draft,
    [switch]$Confirm,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Send As: send from a specific account in your profile (SMTP address or display name;
# list them with outlook-accounts-list.ps1). -Draft saves a draft from that account instead.
# Outward-facing: previews unless -Confirm.
# Usage: .\outlook-send-as.ps1 -Account "work@company.com" -To "a@x.com" -Subject "Hi" -Body "text" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    if ($Draft) { Invoke-NewMail -P $PSBoundParameters -Action "draft" }
    else { Invoke-NewMail -P $PSBoundParameters -Action "send" -Send }
}
catch {
    Stop-WithError $_
}
