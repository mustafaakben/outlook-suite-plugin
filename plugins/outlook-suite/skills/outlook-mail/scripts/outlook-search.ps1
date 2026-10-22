param(
    [string]$Subject = "",
    [string]$From = "",
    [Alias("BodyContains")]
    [string]$Body = "",
    [switch]$FastBody,
    [int]$Days = 30,
    [int]$Limit = 20,
    [string]$Folder = "Inbox",
    [switch]$Json
)

# Outlook Email Search: subject / sender / body text search in one folder (default Inbox).
# Same engine as outlook-find.ps1, which has more filters.
# Usage: .\outlook-search.ps1 -Subject "meeting" -From "John" -Days 7
# Usage: .\outlook-search.ps1 -Body "invoice" -FastBody -Days 365

. (Join-Path $PSScriptRoot "_common.ps1")

if (-not $Subject -and -not $From -and -not $Body) {
    Stop-WithError "No search criteria provided. Use -Subject, -From, or -Body." "Example: .\outlook-search.ps1 -Subject ""meeting"" -From ""John"" -Days 7"
}

try {
    $null = Connect-Outlook
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
    $since = (Get-Date).AddDays(-$Days)
    $result = Find-MailRecords -Folder $target -From $since -Subject $Subject -Sender $From -BodyContains $Body -FastBody:$FastBody -Limit $Limit

    $criteria = @("Last $Days day(s)")
    if ($Subject) { $criteria += "Subject: '$Subject'" }
    if ($From) { $criteria += "From: '$From'" }
    if ($Body) { $criteria += "Body: '$Body'$(if ($FastBody) { ' (indexed)' })" }
    if ($Json) { Write-MailListJson -Result $result -Folder $target -Query @{ filters = $criteria; limit = $Limit } }
    else { Write-MailList -Result $result -Folder $target -Title "SEARCH RESULTS" -Criteria $criteria -Limit $Limit }
}
catch {
    Stop-WithError $_
}
