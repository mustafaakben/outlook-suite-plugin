param(
    [Parameter(Position = 0)]
    [string]$Query = "",
    [ValidateSet("Mail", "Calendar", "Contact", "Task", "Note", "File")]
    [string[]]$Kind = @("Mail", "Calendar", "Contact", "Task", "Note", "File"),
    [int]$Days = 0,
    [string]$From = "",
    [string]$Path = "",
    [int]$Top = 10,
    [switch]$NoResolve,
    [switch]$Json
)

# Outlook Everything: one full-text search across mail, calendar, contacts, tasks, notes AND
# local files, through the Windows Search index (the same index as Outlook's and Explorer's
# search box). Read-only, and fast: no folder walk, the index already has every word.
#
# Each kind is queried separately (up to -Top each), so a flood of files cannot hide mail.
# Outlook hits come back with a real EntryID (the index's short-term id is resolved through
# Outlook), ready for outlook-read-body.ps1, outlook-reply.ps1, calendar or tasks scripts.
# Words are ANDed; "quoted phrases" stay together; trailing * is a prefix match (budg*).
#
# Usage: .\outlook-everything.ps1 "budget transfer"
#        .\outlook-everything.ps1 "AI poster" -Kind Mail,File -Days 90
#        .\outlook-everything.ps1 "syllabus" -Kind File -Path "C:\Users\me\Dropbox\Teaching"
#        .\outlook-everything.ps1 "grant" -Kind Mail -From "jane@example.org" -Top 25 -Json
# Limits: only what Windows Search has indexed (Outlook mail and calendar, plus the folders in
# Indexing Options). Mail moved in the last minutes may not be indexed yet.

. (Join-Path $PSScriptRoot "_common.ps1")

# System.Kind values and the date column used for -Days, per -Kind.
$script:KindMap = [ordered]@{
    Mail     = @{ Where = "System.Kind = 'email' AND System.ItemUrl LIKE 'mapi%'"; Date = "System.Message.DateReceived" }
    Calendar = @{ Where = "System.Kind = 'calendar' AND System.ItemUrl LIKE 'mapi%'"; Date = "System.StartDate" }
    Contact  = @{ Where = "System.Kind = 'contact' AND System.ItemUrl LIKE 'mapi%'"; Date = "System.DateModified" }
    Task     = @{ Where = "System.Kind = 'task' AND System.ItemUrl LIKE 'mapi%'"; Date = "System.DateModified" }
    Note     = @{ Where = "System.Kind = 'note' AND System.ItemUrl LIKE 'mapi%'"; Date = "System.DateModified" }
    File     = @{ Where = "System.ItemUrl LIKE 'file:%'"; Date = "System.DateModified" }
}
$script:Columns = "System.ItemUrl", "System.ItemNameDisplay", "System.Subject", "System.Message.FromName", "System.Message.FromAddress",
"System.Message.DateReceived", "System.StartDate", "System.DateModified", "System.ItemFolderPathDisplay", "System.ItemPathDisplay",
"System.Size", "System.Search.Rank", "System.Message.HasAttachments"

# Words -> CONTAINS('"w1" AND "w2*"'). Single quotes are doubled; double quotes inside a term are dropped.
function ConvertTo-ContainsClause([string]$Text) {
    $terms = @([regex]::Matches($Text, '"([^"]+)"|(\S+)') | ForEach-Object {
            $t = if ($_.Groups[1].Success) { $_.Groups[1].Value } else { $_.Groups[2].Value }
            $t = $t.Replace('"', '').Replace("'", "''").Trim()
            if ($t -and $t -notin @("AND", "OR", "NOT")) { '"' + $t + '"' }
        })
    if ($terms.Count -eq 0) { throw "Give some words to search for." }
    return "CONTAINS('" + ($terms -join " AND ") + "')"
}

# Windows Search dates are UTC; the SQL literal must be UTC too.
function Format-IndexDate([datetime]$Date) { return $Date.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") }
function ConvertFrom-IndexDate($Value) {
    if ($null -eq $Value -or $Value -is [DBNull]) { return $null }
    return [datetime]::SpecifyKind([datetime]$Value, [DateTimeKind]::Utc).ToLocalTime()
}

# mapi16://{SID}/account/0/Folder/<24 chars>: each char is one byte + 0xAC00 of Outlook's
# short-term EntryID, which Namespace.GetItemFromID accepts and expands to the long one.
function Resolve-IndexedItem([string]$Url) {
    $seg = ($Url -split '/')[-1]
    if ($seg.Length -lt 8) { return $null }
    $bytes = New-Object byte[] $seg.Length
    for ($i = 0; $i -lt $seg.Length; $i++) {
        $b = [int]$seg[$i] - 0xAC00
        if ($b -lt 0 -or $b -gt 255) { return $null }
        $bytes[$i] = [byte]$b
    }
    try { return $Namespace.GetItemFromID((ConvertTo-HexId $bytes)) } catch { return $null }
}

function Get-FieldValue($Rs, [string]$Name) {
    $v = $Rs.Fields.Item($Name).Value
    if ($v -is [DBNull]) { return $null }
    return $v
}

try {
    if (-not $Query.Trim()) { Stop-WithError "Give some words to search for." "Example: outlook-everything.ps1 `"budget transfer`" -Kind Mail,File -Days 90" }
    if ($Top -lt 1) { $Top = 1 }
    $contains = ConvertTo-ContainsClause $Query
    $needOutlook = (-not $NoResolve) -and @($Kind | Where-Object { $_ -ne "File" }).Count -gt 0
    if ($needOutlook) { $null = Connect-Outlook }

    $cn = New-Object -ComObject ADODB.Connection
    try { $cn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';") }
    catch { Stop-WithError "Windows Search is not available: $($_.Exception.Message)" "Start the 'Windows Search' service (WSearch), or use outlook-search.ps1 for mail only." }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $groups = [ordered]@{}
    foreach ($k in $script:KindMap.Keys) {
        if ($Kind -notcontains $k) { continue }
        $m = $script:KindMap[$k]
        $where = @($m.Where, $contains)
        if ($Days -gt 0) { $where += "$($m.Date) >= '$(Format-IndexDate (Get-Date).AddDays(-$Days))'" }
        if ($From -and $k -eq "Mail") {
            $f = $From.Replace("'", "''")
            $where += "(System.Message.FromAddress LIKE '%$f%' OR System.Message.FromName LIKE '%$f%')"
        }
        if ($Path -and $k -eq "File") {
            $p = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path.TrimEnd('\')
            $where += "SCOPE='file:$($p.Replace("'", "''"))'"
        }
        $sql = "SELECT TOP $Top $($script:Columns -join ', ') FROM SystemIndex WHERE $($where -join ' AND ') ORDER BY System.Search.Rank DESC"
        $rs = $cn.Execute($sql)
        $rows = New-Object System.Collections.Generic.List[object]
        while (-not $rs.EOF) {
            $url = [string](Get-FieldValue $rs "System.ItemUrl")
            $r = [ordered]@{ Kind = $k }
            if ($k -eq "File") {
                $r.Name = [string](Get-FieldValue $rs "System.ItemNameDisplay")
                $r.Path = [string](Get-FieldValue $rs "System.ItemPathDisplay")
                $r.Modified = ConvertFrom-IndexDate (Get-FieldValue $rs "System.DateModified")
                $size = Get-FieldValue $rs "System.Size"
                $r.SizeKB = if ($null -ne $size) { [math]::Round([double]$size / 1KB, 1) } else { $null }
            }
            else {
                $r.Subject = [string](Get-FieldValue $rs "System.Subject")
                if (-not $r.Subject) { $r.Subject = [string](Get-FieldValue $rs "System.ItemNameDisplay") }
                $r.Folder = ([string](Get-FieldValue $rs "System.ItemFolderPathDisplay")).TrimStart('/')
                if ($k -eq "Mail") {
                    $name = [string](Get-FieldValue $rs "System.Message.FromName"); $addr = [string](Get-FieldValue $rs "System.Message.FromAddress")
                    $r.From = if ($name -and $addr -and $name -ne $addr) { "$name <$addr>" } elseif ($addr) { $addr } else { $name }
                    $r.Date = ConvertFrom-IndexDate (Get-FieldValue $rs "System.Message.DateReceived")
                    $r.HasAttachments = [bool](Get-FieldValue $rs "System.Message.HasAttachments")
                }
                elseif ($k -eq "Calendar") { $r.Start = ConvertFrom-IndexDate (Get-FieldValue $rs "System.StartDate") }
                else { $r.Modified = ConvertFrom-IndexDate (Get-FieldValue $rs "System.DateModified") }
                $r.EntryID = $null
                if (-not $NoResolve) {
                    $item = Resolve-IndexedItem $url
                    if ($item) {
                        $r.EntryID = $item.EntryID
                        # The item is authoritative for the date when the index has none.
                        if ($k -eq "Calendar" -and -not $r.Start) { try { $r.Start = $item.Start } catch { } }
                    }
                }
            }
            $rows.Add($r)
            $rs.MoveNext()
        }
        $rs.Close()
        $groups[$k] = $rows
    }
    $cn.Close()
    $ms = $sw.ElapsedMilliseconds

    if ($Json) {
        $out = [ordered]@{ query = $Query; kinds = @($groups.Keys); days = $Days; elapsedMs = $ms; groups = [ordered]@{} }
        foreach ($k in $groups.Keys) {
            $out.groups[$k] = @($groups[$k] | ForEach-Object {
                    $o = [ordered]@{}
                    foreach ($key in $_.Keys) {
                        $v = $_[$key]
                        $o[$key] = if ($v -is [datetime]) { Format-MailDateIso $v } else { $v }
                    }
                    $o
                })
        }
        Write-Json $out
        exit 0
    }

    $total = 0; foreach ($k in $groups.Keys) { $total += $groups[$k].Count }
    Write-Host "=== EVERYTHING: $Query ===" -ForegroundColor Cyan
    $scope = @("kinds: $($groups.Keys -join ', ')")
    if ($Days -gt 0) { $scope += "last $Days days" }
    if ($From) { $scope += "mail from: $From" }
    if ($Path) { $scope += "files under: $Path" }
    Write-Host "$($scope -join ' | ') | $ms ms" -ForegroundColor Gray
    foreach ($k in $groups.Keys) {
        $rows = $groups[$k]
        Write-Host ""
        Write-Host "$k ($($rows.Count)$(if ($rows.Count -ge $Top) { '+' }))" -ForegroundColor Cyan
        if ($rows.Count -eq 0) { Write-Host "  (none)" -ForegroundColor DarkGray; continue }
        foreach ($r in $rows) {
            switch ($k) {
                "File" { Write-Host "  $(Format-MailDate $r.Modified)  $($r.Name)" -ForegroundColor Yellow; Write-Host "      $($r.Path)" -ForegroundColor Gray }
                "Mail" {
                    $att = if ($r.HasAttachments) { "  [attachments]" } else { "" }
                    Write-Host "  $(Format-MailDate $r.Date)  $($r.Subject)$att" -ForegroundColor Yellow
                    Write-Host "      From: $($r.From)  |  $($r.Folder)" -ForegroundColor Gray
                }
                "Calendar" { Write-Host "  $(Format-MailDate $r.Start)  $($r.Subject)" -ForegroundColor Yellow; Write-Host "      $($r.Folder)" -ForegroundColor Gray }
                default { Write-Host "  $(Format-MailDate $r.Modified)  $($r.Subject)" -ForegroundColor Yellow; Write-Host "      $($r.Folder)" -ForegroundColor Gray }
            }
            if ($k -ne "File") {
                if ($r.EntryID) { Write-Host "      EntryID: $($r.EntryID)" -ForegroundColor DarkGray }
                elseif (-not $NoResolve) { Write-Host "      (not opened: the item may have moved since it was indexed)" -ForegroundColor DarkGray }
            }
        }
    }
    Write-Host ""
    if ($total -eq 0) { Write-Host "Nothing matched. Try fewer words, a prefix like budg*, or drop -Days/-From." -ForegroundColor Gray }
    else { Write-Host "Next: outlook-read-body.ps1 -EntryID <id> | outlook-attachment-read.ps1 -EntryID <id> | outlook-calendar-prep.ps1 -EntryID <id>" -ForegroundColor DarkGray }
}
catch {
    Stop-WithError $_
}
