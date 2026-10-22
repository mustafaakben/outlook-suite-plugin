param(
    [int]$Days = 7,
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [string]$From = "",
    [string]$To = "",
    [string]$Subject = "",
    [string]$BodyContains = "",
    [switch]$FastBody,
    [switch]$UnreadOnly,
    [switch]$ReadOnly,
    [switch]$HasAttachment,
    [switch]$Flagged,
    [ValidateSet("", "Low", "Normal", "High")]
    [string]$Importance = "",
    [string]$Category = "",
    [string]$Folder = "Inbox",
    [switch]$AllStores,
    [int]$Limit = 50,
    [switch]$Detailed,
    [switch]$Json,
    [bool]$StripLinks = $true   # kept for compatibility; list previews always use [URL-00N] references
)

# Outlook Find: list emails in a folder with filters. Returns EntryIDs for every other script.
# Fast path: one Folder.GetTable query; no email is opened unless -Detailed or -BodyContains.
#
# Usage: .\outlook-find.ps1 -Days 1 -UnreadOnly
# Usage: .\outlook-find.ps1 -Subject "budget" -From "jane" -Days 30
# Usage: .\outlook-find.ps1 -Folder "Inbox/Projects" -HasAttachment -Json
# Usage: .\outlook-find.ps1 -BodyContains "invoice" -Days 30            (exact substring, scans candidates)
# Usage: .\outlook-find.ps1 -BodyContains "invoice" -FastBody -Days 365 (content index, word match, fast)

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $targetFolder = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder -AllStores:$AllStores
    $range = Resolve-DateRange -Days $Days -DateFrom $DateFrom -DateTo $DateTo

    $query = @{
        Folder = $targetFolder; From = $range[0]; Subject = $Subject; Sender = $From; Recipient = $To
        BodyContains = $BodyContains; FastBody = $FastBody; UnreadOnly = $UnreadOnly; ReadOnly = $ReadOnly
        HasAttachment = $HasAttachment; Flagged = $Flagged; Importance = $Importance; Category = $Category; Limit = $Limit
    }
    if ($range[1]) { $query.To = $range[1] }
    $result = Find-MailRecords @query
    if ($Detailed) { Add-MailRecordDetails -Namespace $Namespace -Records $result.Records }

    $criteria = @("Since: $(Format-MailDate $range[0])")
    if ($range[1]) { $criteria += "Until: $(Format-MailDate $range[1])" }
    if ($From) { $criteria += "Sender: '$From'" }
    if ($To) { $criteria += "Recipient: '$To'" }
    if ($Subject) { $criteria += "Subject: '$Subject'" }
    if ($BodyContains) { $criteria += "Body: '$BodyContains'$(if ($FastBody) { ' (indexed)' })" }
    if ($UnreadOnly) { $criteria += "Unread only" }
    if ($ReadOnly) { $criteria += "Read only" }
    if ($HasAttachment) { $criteria += "Has attachments" }
    if ($Flagged) { $criteria += "Flagged" }
    if ($Importance) { $criteria += "Importance: $Importance" }
    if ($Category) { $criteria += "Category: '$Category'" }

    if ($Json) {
        Write-MailListJson -Result $result -Folder $targetFolder -Query @{ filters = $criteria; limit = $Limit }
    } else {
        Write-MailList -Result $result -Folder $targetFolder -Title "OUTLOOK FIND" -Criteria $criteria -Limit $Limit
    }
}
catch {
    Stop-WithError $_
}
