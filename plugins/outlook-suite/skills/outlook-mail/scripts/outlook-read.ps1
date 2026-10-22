param(
    [int]$Days = 1,
    [int]$Limit = 20,
    [switch]$UnreadOnly,
    [string]$Priority = "",          # Copilot priority: "High", "High,Normal", "Low"
    [string]$Folder = "Inbox",
    [switch]$Detailed,
    [switch]$Json
)

# Outlook Email Reader: the most recent emails in a folder (default: Inbox, last day).
# For filters use outlook-find.ps1; for a full message use outlook-read-body.ps1.
# Usage: .\outlook-read.ps1 -Days 2 -Limit 50 -UnreadOnly
# Usage: .\outlook-read.ps1 -Days 1 -Priority High        (Copilot's High arrow)

. (Join-Path $PSScriptRoot "_common.ps1")

# Copilot priority for each record (one table read), optionally keeping only $Want levels.
function Add-CopilotPriority($Folder, $Result, $Since, [string[]]$Want, [int]$Limit) {
    $t = $Folder.GetTable('@SQL="urn:schemas:httpmail:datereceived" >= ''' + $Since.ToUniversalTime().ToString('g') + '''')
    $t.Columns.RemoveAll(); [void]$t.Columns.Add($script:Prop.LongTermEntryID); [void]$t.Columns.Add($script:CopilotProp.Score)
    $n = $t.GetRowCount(); $map = @{}
    if ($n -gt 0) {
        $rows = $t.GetArray($n)
        for ($i = 0; $i -lt $n; $i++) { if ($rows[$i, 0] -is [byte[]]) { $map[(ConvertTo-HexId $rows[$i, 0])] = ConvertTo-CopilotPriority $rows[$i, 1] } }
    }
    $records = @(foreach ($r in $Result.Records) { $r.CopilotPriority = [string]$map[$r.EntryID]; if (-not $Want -or $r.CopilotPriority -in $Want) { $r } })
    if ($Want) { $Result.Total = $records.Count; $Result.Records = @($records | Select-Object -First $Limit) }
}

try {
    $want = @($Priority -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($p in $want) { if ($p -notin "High", "Normal", "Low") { throw "-Priority takes High, Normal or Low (got '$p')." } }
    $null = Connect-Outlook
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
    $since = (Get-Date).AddDays(-$Days)
    $result = Find-MailRecords -Folder $target -From $since -UnreadOnly:$UnreadOnly -Limit $(if ($want.Count) { 100000 } else { $Limit })
    Add-CopilotPriority -Folder $target -Result $result -Since $since -Want $want -Limit $Limit
    if ($Detailed) { Add-MailRecordDetails -Namespace $Namespace -Records $result.Records }

    $criteria = @("Last $Days day(s)")
    if ($UnreadOnly) { $criteria += "Unread only" }
    if ($want.Count) { $criteria += "Copilot priority: $($want -join ' or ')" }
    if ($Json) { Write-MailListJson -Result $result -Folder $target -Query @{ filters = $criteria; limit = $Limit } }
    else { Write-MailList -Result $result -Folder $target -Title "EMAILS" -Criteria $criteria -Limit $Limit }
}
catch {
    Stop-WithError $_
}
