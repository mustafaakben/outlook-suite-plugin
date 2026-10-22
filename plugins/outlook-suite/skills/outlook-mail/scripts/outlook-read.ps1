param(
    [int]$Days = 1,
    [int]$Limit = 20,
    [switch]$UnreadOnly,
    [string]$Folder = "Inbox",
    [switch]$Detailed,
    [switch]$Json
)

# Outlook Email Reader: the most recent emails in a folder (default: Inbox, last day).
# For filters use outlook-find.ps1; for a full message use outlook-read-body.ps1.
# Usage: .\outlook-read.ps1 -Days 2 -Limit 50 -UnreadOnly

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
    $since = (Get-Date).AddDays(-$Days)
    $result = Find-MailRecords -Folder $target -From $since -UnreadOnly:$UnreadOnly -Limit $Limit
    if ($Detailed) { Add-MailRecordDetails -Namespace $Namespace -Records $result.Records }

    $criteria = @("Last $Days day(s)")
    if ($UnreadOnly) { $criteria += "Unread only" }
    if ($Json) { Write-MailListJson -Result $result -Folder $target -Query @{ filters = $criteria; limit = $Limit } }
    else { Write-MailList -Result $result -Folder $target -Title "EMAILS" -Criteria $criteria -Limit $Limit }
}
catch {
    Stop-WithError $_
}
