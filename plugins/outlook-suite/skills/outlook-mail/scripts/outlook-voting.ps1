param(
    [Parameter(Mandatory = $true)]
    [string]$To,
    [Parameter(Mandatory = $true)]
    [string]$Subject,
    [string]$Body = "",
    [string]$Options = "Approve;Reject",
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

# Outlook Voting: send an email with voting buttons (Exchange recipients see the buttons).
# -Options: choices separated by ";" (default "Approve;Reject").
# Outward-facing: previews unless -Confirm.
# Usage: .\outlook-voting.ps1 -To "team@x.com" -Subject "Lunch poll" -Body "Where?" -Options "Pizza;Tacos;Sushi" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $p = @{} + $PSBoundParameters
    $p.VotingOptions = $Options
    if (-not $p.ContainsKey("Body")) { $p.Body = "" }
    Invoke-NewMail -P $p -Action "send" -Send
}
catch {
    Stop-WithError $_
}
