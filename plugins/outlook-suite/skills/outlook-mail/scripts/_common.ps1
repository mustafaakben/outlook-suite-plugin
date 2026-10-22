# Shared helpers for the outlook-mail scripts.
# Dot-source at the top of a script:  . (Join-Path $PSScriptRoot "_common.ps1")
#
# Works in two modes:
#   - standalone (pwsh -File script.ps1 ...): connects to Outlook once per call.
#   - inside outlook-host.ps1: reuses the warm Outlook session in $global:OutlookSuite.

Set-StrictMode -Off
. (Join-Path $PSScriptRoot "recipient-resolution.ps1")   # Assert-OutlookRecipients

# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

# Returns the Outlook Application object and sets $script:Outlook / $script:Namespace
# in the calling script. No fixed sleep: Outlook is only waited on when this call
# had to start it.
function Connect-Outlook {
    if ($global:OutlookSuite -and $global:OutlookSuite.App) {
        try {
            $null = $global:OutlookSuite.Namespace.CurrentUser
            Set-Variable -Name Outlook -Scope 1 -Value $global:OutlookSuite.App
            Set-Variable -Name Namespace -Scope 1 -Value $global:OutlookSuite.Namespace
            return $global:OutlookSuite.App
        } catch {
            $global:OutlookSuite.App = $null   # stale session, reconnect below
        }
    }

    $wasRunning = [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    $app = New-Object -ComObject Outlook.Application
    $ns = $app.GetNamespace("MAPI")
    if (-not $wasRunning) {
        # Fresh start: wait until the profile is loaded.
        $deadline = (Get-Date).AddSeconds(30)
        while ((Get-Date) -lt $deadline) {
            try { if ($ns.DefaultStore) { break } } catch { }
            Start-Sleep -Milliseconds 250
        }
    }
    if ($global:OutlookSuite) {
        $global:OutlookSuite.App = $app
        $global:OutlookSuite.Namespace = $ns
    }
    Set-Variable -Name Outlook -Scope 1 -Value $app
    Set-Variable -Name Namespace -Scope 1 -Value $ns
    return $app
}

function Test-InOutlookHost { return [bool]($global:OutlookSuite -and $global:OutlookSuite.InHost) }

# ---------------------------------------------------------------------------
# Errors and output
# ---------------------------------------------------------------------------

# Print "Error: ..." and exit 1. Accepts a string or an ErrorRecord.
function Stop-WithError {
    param($Message, [string]$Hint = "")
    $trace = ""
    if ($Message -is [System.Management.Automation.ErrorRecord]) { $trace = $Message.ScriptStackTrace; $Message = $Message.Exception.Message }
    Write-Host "Error: $Message" -ForegroundColor Red
    if ($env:OUTLOOK_SUITE_DEBUG -eq "1" -and $trace) { Write-Host $trace -ForegroundColor DarkGray }
    if ($Hint) { Write-Host "Hint: $Hint" -ForegroundColor Gray }
    exit 1
}

# Writes an object as compact JSON to stdout (used when a script gets -Json).
function Write-Json {
    param($InputObject, [int]$Depth = 8)
    Write-Output ($InputObject | ConvertTo-Json -Depth $Depth -Compress)
}

# Date shown to agents: local time with UTC offset, e.g. 2026-10-22 09:15 -04:00
function Format-MailDate {
    param($Date)
    if (-not $Date) { return "" }
    $d = [datetime]$Date
    if ($d.Year -ge 4500) { return "" }   # Outlook's "none" date (4501-01-01)
    $dto = [datetimeoffset]::new($d)
    return $dto.ToString("yyyy-MM-dd HH:mm zzz")
}

# Date for JSON: ISO 8601 with offset.
function Format-MailDateIso {
    param($Date)
    if (-not $Date) { return $null }
    $d = [datetime]$Date
    if ($d.Year -ge 4500) { return $null }
    return ([datetimeoffset]::new($d)).ToString("yyyy-MM-ddTHH:mm:sszzz")
}

# Date literal for Items.Restrict / Table.Restrict JET filters. Outlook parses these
# with the same Windows locale as this process, so the short general format is safe here.
function Format-JetDate {
    param([datetime]$Date)
    return $Date.ToString("g")
}

# ---------------------------------------------------------------------------
# Folders
# ---------------------------------------------------------------------------

$script:WellKnownFolders = @{
    "inbox" = 6; "sent" = 5; "sent items" = 5; "drafts" = 16; "deleted" = 3; "deleted items" = 3
    "junk" = 23; "junk email" = 23; "outbox" = 4
}

function Get-FolderCache {
    if ($global:OutlookSuite) {
        if (-not $global:OutlookSuite.FolderCache) { $global:OutlookSuite.FolderCache = @{} }
        return $global:OutlookSuite.FolderCache
    }
    if (-not $script:FolderCacheLocal) { $script:FolderCacheLocal = @{} }
    return $script:FolderCacheLocal
}

# Stores worth searching: the default store first, then other cached mailbox stores.
# Public-folder stores (ExchangeStoreType 2) are slow and skipped unless -AllStores.
function Get-SearchStores {
    param($Namespace, [switch]$AllStores)
    $default = $Namespace.DefaultStore
    $list = New-Object System.Collections.ArrayList
    [void]$list.Add($default)
    foreach ($s in $Namespace.Stores) {
        if ($s.StoreID -eq $default.StoreID) { continue }
        if (-not $AllStores) {
            try { if ($s.ExchangeStoreType -eq 2) { continue } } catch { }
        }
        [void]$list.Add($s)
    }
    return $list
}

function Find-FolderByName {
    param($Root, [string]$Name, [int]$MaxDepth = 6)
    # Breadth-first, so the shallowest match wins.
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue(@($Root, 0))
    while ($queue.Count -gt 0) {
        $pair = $queue.Dequeue(); $f = $pair[0]; $depth = $pair[1]
        foreach ($child in $f.Folders) {
            if ($child.Name -ieq $Name) { return $child }
            if ($depth + 1 -lt $MaxDepth) { $queue.Enqueue(@($child, $depth + 1)) }
        }
    }
    return $null
}

# Resolve a folder from a name or path.
#   "Inbox", "Sent", "Drafts", "Deleted", "Junk", "Outbox"   -> default folders
#   "Inbox/Projects/2026"                                      -> path under the default mailbox
#   "someone@x.edu/Inbox/Sub" or "\\someone@x.edu\Inbox"       -> path in a named mailbox
#   "Projects"                                                 -> first folder with that name (default mailbox first)
# Throws when nothing matches. Never falls back to Inbox.
function Resolve-OutlookFolder {
    param($Namespace, [string]$Path = "Inbox", [switch]$AllStores)
    if (-not $Path) { $Path = "Inbox" }
    $cache = Get-FolderCache
    $key = "$($Path.ToLowerInvariant())|$([bool]$AllStores)"
    if ($cache.ContainsKey($key)) {
        # Use the cached folder only if it still has the requested name (it may have been
        # renamed or moved since; the warm host keeps this cache across calls).
        $last = ($Path.Trim().TrimEnd('\', '/') -split '[\\/]')[-1]
        try {
            $hit = $Namespace.GetFolderFromID($cache[$key][0], $cache[$key][1])
            if ($hit.Name -ieq $last -or $script:WellKnownFolders.ContainsKey($last.ToLowerInvariant())) { return $hit }
        } catch { }
        $cache.Remove($key)
    }

    $clean = $Path.Trim().TrimStart('\', '/').Replace('\', '/')
    $parts = @($clean.Split('/', [System.StringSplitOptions]::RemoveEmptyEntries))
    $folder = $null

    if ($parts.Count -eq 1 -and $script:WellKnownFolders.ContainsKey($parts[0].ToLowerInvariant())) {
        $folder = $Namespace.GetDefaultFolder($script:WellKnownFolders[$parts[0].ToLowerInvariant()])
    } elseif ($parts.Count -ge 1) {
        $stores = Get-SearchStores -Namespace $Namespace -AllStores:$AllStores
        # Named mailbox first: "<store display name>/..."
        $start = $null; $rest = $parts
        foreach ($s in $stores) {
            if ($s.DisplayName -ieq $parts[0]) { $start = $s.GetRootFolder(); $rest = @($parts | Select-Object -Skip 1); break }
        }
        if ($start -and $rest.Count -eq 0) { $folder = $start }
        elseif ($parts.Count -gt 1 -or $start) {
            $roots = if ($start) { @($start) } else { @($stores | ForEach-Object { $_.GetRootFolder() }) }
            foreach ($root in $roots) {
                $f = $root
                foreach ($p in $rest) {
                    $next = $null
                    foreach ($child in $f.Folders) { if ($child.Name -ieq $p) { $next = $child; break } }
                    if (-not $next) { $f = $null; break }
                    $f = $next
                }
                if ($f) { $folder = $f; break }
            }
        } else {
            foreach ($s in $stores) {
                $folder = Find-FolderByName -Root $s.GetRootFolder() -Name $parts[0]
                if ($folder) { break }
            }
        }
    }

    if (-not $folder) {
        throw "Folder '$Path' not found. Use outlook-folders.ps1 to see folder paths (add -AllStores to include public folders)."
    }
    $cache[$key] = @($folder.EntryID, $folder.StoreID)
    return $folder
}

# ---------------------------------------------------------------------------
# Items
# ---------------------------------------------------------------------------

# Get one item by EntryID (preferred) or by -Index (position in a folder's last -Days,
# newest first; volatile, last resort). Throws when not found.
function Get-OutlookItem {
    param(
        $Namespace,
        [string]$EntryID = "",
        [string]$StoreID = "",
        [int]$Index = 0,
        [int]$Days = 7,
        [string]$Folder = "Inbox",
        [string]$DateProperty = "ReceivedTime"
    )
    if ($EntryID) {
        if ($StoreID) {
            try { return $Namespace.GetItemFromID($EntryID, $StoreID) } catch { }
        }
        try { return $Namespace.GetItemFromID($EntryID) } catch { }
        # Not in the default mailbox: try the other mailboxes, so callers never need a StoreID.
        foreach ($s in (Get-SearchStores -Namespace $Namespace)) {
            try { return $Namespace.GetItemFromID($EntryID, $s.StoreID) } catch { }
        }
        throw "No item found with EntryID $($EntryID.Substring(0, [Math]::Min(16, $EntryID.Length)))... It may have been moved (moving can change the EntryID) or deleted. Run outlook-find.ps1 again to get a current EntryID."
    }
    if ($Index -gt 0) {
        Write-Host "Note: -Index is volatile (positions shift as mail arrives). Prefer -EntryID." -ForegroundColor DarkGray
        $f = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
        $since = Format-JetDate (Get-Date).AddDays(-$Days)
        $items = $f.Items.Restrict("[$DateProperty] >= '$since'")
        $items.Sort("[$DateProperty]", $true)
        if ($Index -le $items.Count) { return $items.Item($Index) }
        throw "No item at -Index $Index in '$Folder' for the last $Days days ($($items.Count) items). Use -EntryID from outlook-find.ps1."
    }
    throw "Provide -EntryID (preferred) or -Index to choose an email."
}

# SMTP address of an address entry or the sender, resolving Exchange (EX) addresses.
$script:PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
$script:PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"

function Get-SenderSmtp {
    param($Item)
    try {
        $addr = $Item.SenderEmailAddress
        if ($addr -and $addr -notmatch '^/O=') { return $addr }
        try { $smtp = $Item.PropertyAccessor.GetProperty($script:PR_SENDER_SMTP); if ($smtp) { return $smtp } } catch { }
        try { $ex = $Item.Sender.GetExchangeUser(); if ($ex) { return $ex.PrimarySmtpAddress } } catch { }
        return $addr
    } catch { return "" }
}

function Get-RecipientList {
    param($Item, [int]$Type)   # 1 = To, 2 = CC, 3 = BCC
    $out = @()
    try {
        foreach ($r in $Item.Recipients) {
            if ($r.Type -ne $Type) { continue }
            $addr = $r.Address
            if (-not $addr -or $addr -match '^/O=') {
                try { $addr = $r.PropertyAccessor.GetProperty($script:PR_SMTP_ADDRESS) } catch { }
            }
            $out += if ($addr -and $addr -ne $r.Name) { "$($r.Name) <$addr>" } else { $r.Name }
        }
    } catch { }
    return $out
}

$script:ImportanceNames = @{ 0 = "Low"; 1 = "Normal"; 2 = "High" }

$script:CategoryColors = @("None", "Red", "Orange", "Peach", "Yellow", "Green", "Teal", "Olive", "Blue", "Purple",
    "Maroon", "Steel", "DarkSteel", "Gray", "DarkGray", "Black", "DarkRed", "DarkOrange", "DarkPeach", "DarkYellow",
    "DarkGreen", "DarkTeal", "DarkOlive", "DarkBlue", "DarkPurple", "DarkMaroon")

# Color given as a name ("Green") or an index ("5") -> index 0..25.
function Resolve-CategoryColor([string]$Color) {
    if ($Color -match '^\d+$' -and [int]$Color -le 25) { return [int]$Color }
    for ($i = 0; $i -lt $script:CategoryColors.Count; $i++) { if ($script:CategoryColors[$i] -ieq $Color) { return $i } }
    throw "Unknown color '$Color'. Use a name ($($script:CategoryColors -join ', ')) or 0-25."
}

function Get-OutlookCategory($Namespace, [string]$Name) {
    foreach ($c in $Namespace.Categories) { if ($c.Name -ieq $Name) { return $c } }
    return $null
}

# A full description of one item, used by single-item views and action results.
function ConvertTo-MailRecord {
    param($Item, [switch]$WithRecipients, [switch]$WithBody, [string]$Links = "domain")
    $rec = [ordered]@{
        EntryID        = $Item.EntryID
        StoreID        = ""
        ConversationID = ""
        Folder         = ""
        Subject        = $Item.Subject
        From           = ""
        Date           = $null
        Unread         = $null
        Importance     = ""
        Flag           = ""
        Categories     = $Item.Categories
        Attachments    = @()
    }
    try { $rec.StoreID = $Item.Parent.StoreID } catch { }
    try { $rec.ConversationID = $Item.ConversationID } catch { }
    try { $rec.Folder = $Item.Parent.FolderPath } catch { }
    try {
        $name = $Item.SenderName; $smtp = Get-SenderSmtp $Item
        $rec.From = if ($smtp -and $smtp -ne $name) { "$name <$smtp>" } else { $name }
    } catch { }
    $date = $null
    foreach ($p in "ReceivedTime", "SentOn", "LastModificationTime") {
        try { $v = $Item.$p; if ($v -and ([datetime]$v).Year -lt 4500) { $date = $v; break } } catch { }
    }
    $rec.Date = $date
    try { $rec.Unread = [bool]$Item.UnRead } catch { }
    try { $rec.Importance = $script:ImportanceNames[[int]$Item.Importance] } catch { }
    try {
        $rec.Flag = switch ([int]$Item.FlagStatus) { 2 { "Flagged" } 1 { "Complete" } default { "" } }
        if ($rec.Flag -and $Item.TaskDueDate -and ([datetime]$Item.TaskDueDate).Year -lt 4500) { $rec.Flag += " (due $(Format-MailDate $Item.TaskDueDate))" }
    } catch { }
    try {
        for ($i = 1; $i -le $Item.Attachments.Count; $i++) { $rec.Attachments += $Item.Attachments.Item($i).FileName }
    } catch { }
    if ($WithRecipients) {
        $rec.To = @(Get-RecipientList $Item 1)
        $rec.CC = @(Get-RecipientList $Item 2)
        $rec.BCC = @(Get-RecipientList $Item 3)
    }
    if ($WithBody) {
        $html = ""; try { $html = $Item.HTMLBody } catch { }
        $res = Format-MailLinks -Text $Item.Body -Html $html -Mode $Links
        $rec.Body = $res.Text
        $rec.Links = $res.Links
    }
    return $rec
}

# Text rendering of a record. Shows every field an agent needs to act on the item.
function Write-MailRecord {
    param($Record, [string]$Header = "", [switch]$LinkTable)
    if ($Header) { Write-Host $Header -ForegroundColor Cyan }
    Write-Host "  EntryID: $($Record.EntryID)"
    # StoreID is long and not needed (lookups search every mailbox); it stays in -Json output.
    if ($Record.ConversationID) { Write-Host "  ConversationID: $($Record.ConversationID)" -ForegroundColor DarkGray }
    if ($Record.Folder) { Write-Host "  Folder: $($Record.Folder)" -ForegroundColor DarkGray }
    Write-Host "  Subject: $($Record.Subject)" -ForegroundColor Yellow
    if ($Record.From) { Write-Host "  From: $($Record.From)" }
    if ($Record.Contains("To") -and $Record.To) { Write-Host "  To: $($Record.To -join '; ')" }
    if ($Record.Contains("CC") -and $Record.CC) { Write-Host "  CC: $($Record.CC -join '; ')" }
    if ($Record.Contains("BCC") -and $Record.BCC) { Write-Host "  BCC: $($Record.BCC -join '; ')" }
    if ($Record.Date) { Write-Host "  Date: $(Format-MailDate $Record.Date)" }
    if ($null -ne $Record.Unread) { Write-Host "  Status: $(if ($Record.Unread) { 'UNREAD' } else { 'read' })" }
    if ($Record.Importance -and $Record.Importance -ne "Normal") { Write-Host "  Importance: $($Record.Importance)" -ForegroundColor Magenta }
    if ($Record.Flag) { Write-Host "  Flag: $($Record.Flag)" -ForegroundColor Red }
    if ($Record.Categories) { Write-Host "  Categories: $($Record.Categories)" -ForegroundColor Green }
    if ($Record.Attachments -and $Record.Attachments.Count -gt 0) { Write-Host "  Attachments ($($Record.Attachments.Count)): $($Record.Attachments -join ', ')" }
    if ($Record.Contains("Preview") -and $Record.Preview) { Write-Host "  Preview: $($Record.Preview)" -ForegroundColor Gray }
    if ($Record.Contains("Body")) {
        Write-Host "  Body:"
        Write-Host $Record.Body
        if ($LinkTable -and $Record.Links -and $Record.Links.Count -gt 0) {
            Write-Host "  Links (full URLs: outlook-links.ps1 -EntryID <EntryID> [-Ref URL-001]):" -ForegroundColor DarkGray
            foreach ($l in $Record.Links) {
                $t = if ($l.Text) { "  ""$($l.Text)""" } else { "" }
                Write-Host "    $($l.Ref)  $($l.Host)$t" -ForegroundColor DarkGray
            }
        }
    }
}

# Standard result of an action on one item: title, the item (EntryID first), what changed.
# -Changes is an ordered hashtable of label -> value. With -Json prints one object instead.
function Write-ActionResult {
    param([string]$Title, $Item, $Changes = [ordered]@{}, [switch]$Json, [string]$Action = "")
    $rec = ConvertTo-MailRecord $Item
    if ($Json) {
        $o = ConvertTo-JsonRecord $rec
        $o.Action = $Action; $o.Result = $Title
        foreach ($k in $Changes.Keys) { $o[$k] = $Changes[$k] }
        Write-Json $o
        return
    }
    Write-Host "=== $Title ===" -ForegroundColor Green
    Write-Host "  EntryID: $($rec.EntryID)"
    Write-Host "  Subject: $($rec.Subject)" -ForegroundColor Yellow
    if ($rec.From) { Write-Host "  From: $($rec.From)" }
    if ($rec.Date) { Write-Host "  Date: $(Format-MailDate $rec.Date)" }
    if ($rec.Folder) { Write-Host "  Folder: $($rec.Folder)" -ForegroundColor DarkGray }
    foreach ($k in $Changes.Keys) { Write-Host "  ${k}: $($Changes[$k])" -ForegroundColor Cyan }
}

# Shows what an irreversible action would do, then asks Approve-Action. Returns $true to proceed.
function Confirm-ItemAction {
    param([string]$Action, [string]$Title, $Item, [bool]$Confirmed, $Details = [ordered]@{}, [string]$What = "")
    if ($Confirmed -or -not (Test-ConfirmRequired -Action $Action)) { return $true }
    $rec = ConvertTo-MailRecord $Item
    Write-Host "=== $Title (PREVIEW) ===" -ForegroundColor Yellow
    Write-Host "  EntryID: $($rec.EntryID)"
    Write-Host "  Subject: $($rec.Subject)" -ForegroundColor Yellow
    if ($rec.From) { Write-Host "  From: $($rec.From)" }
    if ($rec.Date) { Write-Host "  Date: $(Format-MailDate $rec.Date)" }
    if ($rec.Folder) { Write-Host "  Folder: $($rec.Folder)" -ForegroundColor DarkGray }
    foreach ($k in $Details.Keys) { Write-Host "  ${k}: $($Details[$k])" -ForegroundColor Cyan }
    return (Approve-Action -Action $Action -Confirmed $false -What $(if ($What) { $What } else { $Title.ToLower() }))
}

# JSON-friendly copy of a record (dates as ISO 8601).
function ConvertTo-JsonRecord {
    param($Record)
    $o = [ordered]@{}
    foreach ($k in $Record.Keys) {
        $o[$k] = if ($k -eq "Date") { Format-MailDateIso $Record[$k] } else { $Record[$k] }
    }
    return $o
}

# ---------------------------------------------------------------------------
# Fast listing with Folder.GetTable
# ---------------------------------------------------------------------------

$script:Prop = @{
    LongTermEntryID = "http://schemas.microsoft.com/mapi/proptag/0x66700102"   # same value as item.EntryID
    ConversationID  = "http://schemas.microsoft.com/mapi/proptag/0x30130102"
    MessageFlags    = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"   # bit 1 = read
    HasAttach       = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"
    FlagStatus      = "http://schemas.microsoft.com/mapi/proptag/0x10900003"
    SenderSmtp      = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
    Preview         = "urn:schemas:httpmail:textdescription"                   # first 255 chars of the body
}

function ConvertTo-DaslText([string]$s) { return $s.Replace("'", "''") }

# Build a DASL filter (all comparisons by namespace; dates in UTC as DASL requires).
function New-MailDaslFilter {
    param(
        [datetime]$From, [datetime]$To,
        [string]$Subject, [string]$Sender, [string]$Recipient,
        [switch]$UnreadOnly, [switch]$ReadOnly, [switch]$HasAttachment, [switch]$Flagged,
        [string]$Importance, [string]$Category, [string]$BodyIndexed
    )
    $q = @()
    $d = '"urn:schemas:httpmail:datereceived"'
    if ($From) { $q += "$d >= '$($From.ToUniversalTime().ToString('g'))'" }
    if ($To) { $q += "$d <= '$($To.ToUniversalTime().ToString('g'))'" }
    if ($Subject) { $q += """urn:schemas:httpmail:subject"" LIKE '%$(ConvertTo-DaslText $Subject)%'" }
    if ($Sender) {
        $s = ConvertTo-DaslText $Sender
        $q += "(""urn:schemas:httpmail:fromname"" LIKE '%$s%' OR ""$($script:Prop.SenderSmtp)"" LIKE '%$s%' OR ""urn:schemas:httpmail:fromemail"" LIKE '%$s%')"
    }
    if ($Recipient) { $q += """urn:schemas:httpmail:displayto"" LIKE '%$(ConvertTo-DaslText $Recipient)%'" }
    if ($UnreadOnly) { $q += """urn:schemas:httpmail:read"" = 0" }
    if ($ReadOnly) { $q += """urn:schemas:httpmail:read"" = 1" }
    if ($HasAttachment) { $q += """urn:schemas:httpmail:hasattachment"" = 1" }
    if ($Flagged) { $q += """$($script:Prop.FlagStatus)"" = 2" }
    if ($Importance) {
        $map = @{ "Low" = 0; "Normal" = 1; "High" = 2 }
        $q += """urn:schemas:httpmail:importance"" = $($map[$Importance])"
    }
    if ($Category) { $q += """urn:schemas-microsoft-com:office:office#Keywords"" LIKE '%$(ConvertTo-DaslText $Category)%'" }
    if ($BodyIndexed) { $q += """urn:schemas:httpmail:textdescription"" ci_phrasematch '$(ConvertTo-DaslText $BodyIndexed)'" }
    if ($q.Count -eq 0) { return "" }
    return "@SQL=" + ($q -join " AND ")
}

# Rows of a folder as list records, newest first. Reads only table columns; no item is opened.
function Get-MailTableRecords {
    param($Folder, [string]$Filter = "", [int]$Limit = 50, [string]$SortBy = "ReceivedTime")
    $table = if ($Filter) { $Folder.GetTable($Filter) } else { $Folder.GetTable() }
    $table.Columns.RemoveAll()
    $cols = @("EntryID", "Subject", "SenderName", $script:Prop.SenderSmtp, "To", "ReceivedTime",
        $script:Prop.MessageFlags, $script:Prop.HasAttach, "Importance", "Categories", $script:Prop.FlagStatus,
        $script:Prop.Preview, $script:Prop.LongTermEntryID, $script:Prop.ConversationID)
    foreach ($c in $cols) { [void]$table.Columns.Add($c) }
    $table.Sort($SortBy, $true)
    $total = $table.GetRowCount()
    $n = [Math]::Min($Limit, $total)
    $rows = $null
    if ($n -gt 0) { $rows = $table.GetArray($n) }   # direct assignment keeps the 2-D array (an if-expression would flatten it)
    $store = $Folder.StoreID; $path = $Folder.FolderPath
    $list = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $n; $i++) {
        $eid = $rows[$i, 12]; $eid = if ($eid -is [byte[]]) { [Convert]::ToHexString($eid) } else { [string]$rows[$i, 0] }
        $conv = $rows[$i, 13]; $conv = if ($conv -is [byte[]]) { [Convert]::ToHexString($conv) } else { "" }
        $name = [string]$rows[$i, 2]; $smtp = [string]$rows[$i, 3]
        $flags = $rows[$i, 6]
        $flagStatus = [int]$rows[$i, 10]
        $rec = [ordered]@{
            EntryID        = $eid
            StoreID        = $store
            ConversationID = $conv
            Folder         = $path
            Subject        = [string]$rows[$i, 1]
            From           = $(if ($smtp -and $smtp -ne $name) { "$name <$smtp>" } else { $name })
            ToLine         = [string]$rows[$i, 4]
            Date           = $rows[$i, 5]
            Unread         = $(if ($null -ne $flags) { -not ([int]$flags -band 1) } else { $null })
            Importance     = $script:ImportanceNames[[int]$rows[$i, 8]]
            Flag           = $(if ($flagStatus -eq 2) { "Flagged" } elseif ($flagStatus -eq 1) { "Complete" } else { "" })
            Categories     = [string]$rows[$i, 9]
            HasAttachments = [bool]$rows[$i, 7]
            Preview        = (Format-Preview ([string]$rows[$i, 11]) 200)
        }
        $list.Add($rec)
    }
    return [pscustomobject]@{ Total = $total; Records = $list }
}

# The shared search engine behind outlook-find / outlook-read / outlook-search.
# Everything except substring body search runs as one indexed Table query.
# -BodyContains keeps exact substring semantics (nothing missed): it scans only the
# items that already passed the other filters. -FastBody uses the content index
# instead (word match, much faster, may miss mail not yet indexed).
function Find-MailRecords {
    param(
        $Folder,
        [datetime]$From, [datetime]$To,
        [string]$Subject, [string]$Sender, [string]$Recipient, [string]$BodyContains,
        [switch]$FastBody,
        [switch]$UnreadOnly, [switch]$ReadOnly, [switch]$HasAttachment, [switch]$Flagged,
        [string]$Importance, [string]$Category,
        [int]$Limit = 50, [int]$MaxScan = 5000
    )
    $common = @{ From = $From; To = $To; Subject = $Subject; Sender = $Sender; Recipient = $Recipient
        UnreadOnly = $UnreadOnly; ReadOnly = $ReadOnly; HasAttachment = $HasAttachment; Flagged = $Flagged
        Importance = $Importance; Category = $Category }
    if (-not $From) { $common.Remove("From") }
    if (-not $To) { $common.Remove("To") }

    if (-not $BodyContains -or $FastBody) {
        if ($BodyContains) { $common.BodyIndexed = $BodyContains }
        $filter = New-MailDaslFilter @common
        return Get-MailTableRecords -Folder $Folder -Filter $filter -Limit $Limit
    }

    # Exact substring body search over the pre-filtered candidates.
    $filter = New-MailDaslFilter @common
    $items = $Folder.Items
    if ($filter) { $items = $Folder.Items.Restrict($filter) }
    $items.Sort("[ReceivedTime]", $true)
    $hits = New-Object System.Collections.Generic.List[string]
    $scanned = 0
    foreach ($it in $items) {
        $scanned++
        if ($scanned -gt $MaxScan) { break }
        try { if ($it.Body -like "*$BodyContains*") { $hits.Add($it.EntryID) } } catch { }
        if ($hits.Count -ge $Limit) { break }
    }
    if ($hits.Count -eq 0) { return [pscustomobject]@{ Total = 0; Records = @(); Scanned = $scanned } }
    # Re-read the hits through the table so records have the same shape as other searches.
    $all = Get-MailTableRecords -Folder $Folder -Filter $filter -Limit 100000
    $keep = @{}; foreach ($h in $hits) { $keep[$h] = $true }
    $records = @($all.Records | Where-Object { $keep.ContainsKey($_.EntryID) })
    return [pscustomobject]@{ Total = $records.Count; Records = $records; Scanned = $scanned }
}

# Adds recipients and attachment names by opening each item (slower; for -Detailed).
function Add-MailRecordDetails {
    param($Namespace, $Records)
    foreach ($r in $Records) {
        try {
            $it = $Namespace.GetItemFromID($r.EntryID, $r.StoreID)
            $r.To = @(Get-RecipientList $it 1)
            $r.CC = @(Get-RecipientList $it 2)
            $names = @(); for ($i = 1; $i -le $it.Attachments.Count; $i++) { $names += $it.Attachments.Item($i).FileName }
            $r.Attachments = $names
        } catch { }
    }
}

# Text rendering of list results. Folder and StoreID are printed once, not per row.
function Write-MailList {
    param($Result, $Folder, [string]$Title = "OUTLOOK FIND", [string[]]$Criteria = @(), [int]$Limit = 50)
    Write-Host "=== $Title ===" -ForegroundColor Cyan
    Write-Host "Folder: $($Folder.FolderPath)"
    if ($Criteria.Count -gt 0) { Write-Host "Filters: $($Criteria -join ' | ')" -ForegroundColor Gray }
    Write-Host ""
    $n = 0
    foreach ($r in $Result.Records) {
        $n++
        Write-Host "--- #$n ---" -ForegroundColor Cyan
        Write-Host "  EntryID: $($r.EntryID)"
        if ($r.ConversationID) { Write-Host "  ConversationID: $($r.ConversationID)" -ForegroundColor DarkGray }
        Write-Host "  Subject: $($r.Subject)" -ForegroundColor Yellow
        Write-Host "  From: $($r.From)"
        if ($r.Contains("To") -and $r.To) { Write-Host "  To: $($r.To -join '; ')" } elseif ($r.ToLine) { Write-Host "  To: $($r.ToLine)" }
        if ($r.Contains("CC") -and $r.CC) { Write-Host "  CC: $($r.CC -join '; ')" }
        Write-Host "  Date: $(Format-MailDate $r.Date)"
        Write-Host "  Status: $(if ($r.Unread) { 'UNREAD' } else { 'read' })"
        if ($r.Importance -and $r.Importance -ne "Normal") { Write-Host "  Importance: $($r.Importance)" -ForegroundColor Magenta }
        if ($r.Flag) { Write-Host "  Flag: $($r.Flag)" -ForegroundColor Red }
        if ($r.Categories) { Write-Host "  Categories: $($r.Categories)" -ForegroundColor Green }
        if ($r.Contains("Attachments") -and $r.Attachments) { Write-Host "  Attachments ($($r.Attachments.Count)): $($r.Attachments -join ', ')" }
        elseif ($r.HasAttachments) { Write-Host "  Attachments: yes (names: outlook-attachment-save.ps1 -EntryID <id> -ListOnly)" }
        if ($r.Preview) { Write-Host "  Preview: $($r.Preview)" -ForegroundColor Gray }
        Write-Host ""
    }
    if ($n -eq 0) {
        Write-Host "No matching emails found. Try a larger -Days, fewer filters, or another -Folder (outlook-folders.ps1)." -ForegroundColor Gray
    } else {
        $more = if ($Result.Total -gt $n) { " of $($Result.Total) matching; raise -Limit for more" } else { "" }
        Write-Host "=== Showing $n$more ===" -ForegroundColor Cyan
        Write-Host "Next: outlook-read-body.ps1 -EntryID <EntryID> | outlook-conversation.ps1 -EntryID <EntryID> | outlook-reply.ps1 -EntryID <EntryID> ..." -ForegroundColor DarkGray
    }
}

# JSON rendering of list results.
function Write-MailListJson {
    param($Result, $Folder, [hashtable]$Query = @{})
    $records = @(foreach ($r in $Result.Records) {
        $o = ConvertTo-JsonRecord $r
        $o.Remove("Folder"); $o.Remove("StoreID")
        $o
    })
    Write-Json ([ordered]@{ folder = $Folder.FolderPath; storeId = $Folder.StoreID; total = $Result.Total; count = $records.Count; query = $Query; items = $records })
}

# Resolve -Days / -DateFrom / -DateTo into a [From, To] pair (To may be $null = now).
function Resolve-DateRange {
    param([int]$Days = 7, $DateFrom, $DateTo)
    if ($DateFrom -and $DateTo) { return @([datetime]$DateFrom, ([datetime]$DateTo).Date.AddDays(1).AddMinutes(-1)) }
    if ($DateFrom) { return @([datetime]$DateFrom, $null) }
    if ($DateTo) { $t = ([datetime]$DateTo).Date.AddDays(1).AddMinutes(-1); return @($t.AddDays(-$Days), $t) }
    return @((Get-Date).AddDays(-$Days), $null)
}

# ---------------------------------------------------------------------------
# Links: citation-style references instead of long URLs
# ---------------------------------------------------------------------------

# Unwrap Microsoft Safe Links (and similar ?url= redirectors) to the real target.
function Resolve-RealUrl {
    param([string]$Url)
    if ($Url -match '^https?://[^/]*safelinks\.protection\.outlook\.com/.*?[?&]url=([^&]+)') {
        try { return [uri]::UnescapeDataString($Matches[1]) } catch { }
    }
    return $Url
}

function Get-UrlHost {
    param([string]$Url)
    try {
        $u = if ($Url -match '^(?i)www\.') { "http://$Url" } else { $Url }
        return ([uri]$u).Host -replace '^www\.', ''
    } catch { return ($Url -replace '^(?i)https?://', '').Split('/')[0] }
}

# Replace URLs in a plain-text body with [URL-001] style references.
#   Mode "domain" (default): references inline, plus a list of {Ref, Host, Text, Url}.
#   Mode "full": text unchanged (Safe Links still decoded), list still returned.
#   Mode "none": URLs removed, no list.
# Numbering follows first appearance in the text, so the same message always gives the
# same numbers; outlook-links.ps1 relies on that to return full URLs later.
function Format-MailLinks {
    param([string]$Text, [string]$Html = "", [ValidateSet("domain", "full", "none")][string]$Mode = "domain")
    $result = [ordered]@{ Text = $Text; Links = @() }
    if (-not $Text) { return $result }

    # Anchor text from the HTML body, keyed by real URL.
    $anchorText = @{}
    if ($Html) {
        foreach ($m in [regex]::Matches($Html, '(?is)<a\s[^>]*?href\s*=\s*["'']([^"'']+)["''][^>]*>(.*?)</a>')) {
            $href = Resolve-RealUrl ([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value))
            $txt = [System.Net.WebUtility]::HtmlDecode(($m.Groups[2].Value -replace '<[^>]+>', '')) -replace '\s+', ' '
            $txt = $txt.Trim()
            if ($txt -and -not $anchorText.ContainsKey($href) -and $txt -ne $href -and $txt.Length -le 80) { $anchorText[$href] = $txt }
        }
    }

    # mailto: links keep the address (short and useful), drop the scheme.
    $Text = [regex]::Replace($Text, '(?i)<?mailto:([^\s>]+)>?', '<$1>')

    $refs = [ordered]@{}
    $pattern = '(?i)<?((?:https?://|www\.)[^\s<>"''\]]+)>?'
    $evaluator = {
        param($m)
        $raw = $m.Groups[1].Value.TrimEnd('.', ',', ';', ':', ')', '!', '?')
        $trail = $m.Groups[1].Value.Substring($raw.Length)
        $real = Resolve-RealUrl $raw
        if ($Mode -eq "none") { return $trail }
        if ($Mode -eq "full") { return "<$real>$trail" }
        if (-not $refs.Contains($real)) { $refs[$real] = "URL-{0:000}" -f ($refs.Count + 1) }
        return "[$($refs[$real])]$trail"
    }

    if ($Mode -eq "full") {
        # Collect references for the list even in full mode.
        foreach ($m in [regex]::Matches($Text, $pattern)) {
            $raw = $m.Groups[1].Value.TrimEnd('.', ',', ';', ':', ')', '!', '?')
            $real = Resolve-RealUrl $raw
            if (-not $refs.Contains($real)) { $refs[$real] = "URL-{0:000}" -f ($refs.Count + 1) }
        }
    }
    $result.Text = [regex]::Replace($Text, $pattern, $evaluator)
    if ($Mode -ne "none") {
        $result.Links = @(foreach ($url in $refs.Keys) {
            [ordered]@{ Ref = $refs[$url]; Host = (Get-UrlHost $url); Text = $anchorText[$url]; Url = $url }
        })
    }
    return $result
}

# Short single-line preview for list views (links as [URL-00N], no table).
function Format-Preview {
    param([string]$Text, [int]$Length = 200)
    if (-not $Text) { return "" }
    $t = ($Text -replace '\s+', ' ').Trim()
    $t = (Format-MailLinks -Text $t -Mode domain).Text
    if ($t.Length -gt $Length) { $t = $t.Substring(0, $Length) + "..." }
    return $t
}

# ---------------------------------------------------------------------------
# Rules
# ---------------------------------------------------------------------------

$script:RuleConditionNames = @{
    0 = "Unknown"; 1 = "From"; 2 = "Subject contains"; 3 = "Account"; 4 = "Sent only to me"; 5 = "My name in To"
    6 = "Importance"; 7 = "Sensitivity"; 8 = "Flagged for action"; 9 = "My name in Cc"; 10 = "My name in To or Cc"
    11 = "My name not in To"; 12 = "Sent to"; 13 = "Body contains"; 14 = "Body or subject contains"
    15 = "Header contains"; 16 = "Recipient address contains"; 17 = "Sender address contains"; 18 = "Category"
    19 = "Out-of-office message"; 20 = "Has attachment"; 21 = "Size range"; 22 = "Date range"; 23 = "Form name"
    24 = "Document property"; 25 = "Sender in address book"; 26 = "Meeting invite or update"; 27 = "On this computer only"
    28 = "On another computer"; 29 = "Any category"; 30 = "From RSS feed"; 31 = "From any RSS feed"
}
$script:RuleActionNames = @{
    0 = "Unknown"; 1 = "Move to folder"; 2 = "Assign category"; 3 = "Delete"; 4 = "Delete permanently"; 5 = "Copy to folder"
    6 = "Forward"; 7 = "Forward as attachment"; 8 = "Redirect"; 9 = "Server reply"; 10 = "Reply with template"
    11 = "Flag for action in days"; 12 = "Flag color"; 13 = "Clear flag"; 14 = "Set importance"; 15 = "Set sensitivity"
    16 = "Print"; 17 = "Play sound"; 18 = "Start application"; 19 = "Mark read"; 20 = "Run script"; 21 = "Stop processing more rules"
    22 = "Custom action"; 23 = "New item alert"; 24 = "Desktop alert"; 25 = "Request read receipt"; 26 = "Request delivery receipt"
    27 = "Cc the message"; 28 = "Defer delivery"; 30 = "Clear categories"; 41 = "Mark as task"
}

# The store whose rules to use: -Mailbox (display name or SMTP) or the default mailbox.
function Get-RuleStore {
    param($Namespace, [string]$Mailbox = "")
    if (-not $Mailbox) { return $Namespace.DefaultStore }
    foreach ($s in $Namespace.Stores) { if ($s.DisplayName -ieq $Mailbox) { return $s } }
    throw "Mailbox '$Mailbox' not found. Mailboxes: $(@(foreach ($s in $Namespace.Stores) { $s.DisplayName }) -join '; ')"
}

function Get-StoreRules {
    param($Store)
    # The unary comma returns the Rules collection itself; without it PowerShell would unroll it
    # (an empty collection would come back as $null, and Create/Save would be lost).
    try { return , $Store.GetRules() }
    catch {
        $m = $_.Exception.Message
        if ($m -match "conflict") {
            throw "Outlook reports that rules in mailbox '$($Store.DisplayName)' are in conflict (the copy in Outlook and the copy on the server differ), so no rule can be read or changed by script until that is resolved once in Outlook: File > Manage Rules & Alerts, then choose which copy to keep."
        }
        if ($m -match "does not support rules|could not be found") { throw "Mailbox '$($Store.DisplayName)' does not support rules." }
        throw
    }
}

function Get-RuleDetail($obj, [int]$type, [bool]$isAction) {
    $d = @()
    try { if ($obj.Text) { $d += "'" + (@($obj.Text) -join "' or '") + "'" } } catch { }
    try { if ($obj.Recipients -and $obj.Recipients.Count) { $d += (@(foreach ($r in $obj.Recipients) { if ($r.Address -and $r.Address -notmatch '^/O=') { $r.Address } else { $r.Name } }) -join "; ") } } catch { }
    try { if ($obj.Categories) { $d += (@($obj.Categories) -join ", ") } } catch { }
    try { if ($isAction -and $obj.Folder) { $d += $obj.Folder.FolderPath } } catch { }
    try { if ($type -eq 6 -and -not $isAction) { $d += $script:ImportanceNames[[int]$obj.Importance] } } catch { }
    try { if ($type -eq 3 -and -not $isAction -and $obj.Account) { $d += $obj.Account.DisplayName } } catch { }
    try { if ($type -eq 25 -and $obj.AddressList) { $d += $obj.AddressList.Name } } catch { }
    try { if ($type -eq 23 -and $obj.FormName) { $d += (@($obj.FormName) -join ", ") } } catch { }
    try { if ($type -eq 17 -and $isAction -and $obj.FilePath) { $d += $obj.FilePath } } catch { }
    return ($d -join " | ")
}

# A rule as an ordered record: name, on/off, order, type, conditions, exceptions, actions.
function ConvertTo-RuleRecord {
    param($Rule)
    $conds = @(); $excs = @(); $acts = @()
    for ($i = 1; $i -le $Rule.Conditions.Count; $i++) {
        $c = $Rule.Conditions.Item($i)
        if ($c.Enabled) { $t = [int]$c.ConditionType; $conds += [ordered]@{ Type = $script:RuleConditionNames[$t]; Value = (Get-RuleDetail $c $t $false) } }
    }
    for ($i = 1; $i -le $Rule.Exceptions.Count; $i++) {
        $c = $Rule.Exceptions.Item($i)
        if ($c.Enabled) { $t = [int]$c.ConditionType; $excs += [ordered]@{ Type = $script:RuleConditionNames[$t]; Value = (Get-RuleDetail $c $t $false) } }
    }
    for ($i = 1; $i -le $Rule.Actions.Count; $i++) {
        $a = $Rule.Actions.Item($i)
        if ($a.Enabled) { $t = [int]$a.ActionType; $acts += [ordered]@{ Type = $script:RuleActionNames[$t]; Value = (Get-RuleDetail $a $t $true) } }
    }
    return [ordered]@{
        Order = [int]$Rule.ExecutionOrder; Name = $Rule.Name; Enabled = [bool]$Rule.Enabled
        Kind = $(if ([int]$Rule.RuleType -eq 1) { "send" } else { "receive" }); ClientOnly = [bool]$Rule.IsLocalRule
        Conditions = $conds; Exceptions = $excs; Actions = $acts
    }
}

function Write-RuleRecord {
    param($R, [switch]$Detailed)
    $state = if ($R.Enabled) { "ON " } else { "OFF" }
    $extra = @(); if ($R.Kind -eq "send") { $extra += "send rule" }; if ($R.ClientOnly) { $extra += "client-only" }
    Write-Host ("{0,3}. [{1}] {2}{3}" -f $R.Order, $state, $R.Name, $(if ($extra) { "  ($($extra -join ', '))" })) -ForegroundColor $(if ($R.Enabled) { "Yellow" } else { "Gray" })
    if ($Detailed) {
        foreach ($c in $R.Conditions) { Write-Host "       if:     $($c.Type)$(if ($c.Value) { ": $($c.Value)" })" }
        foreach ($c in $R.Exceptions) { Write-Host "       except: $($c.Type)$(if ($c.Value) { ": $($c.Value)" })" }
        foreach ($a in $R.Actions) { Write-Host "       then:   $($a.Type)$(if ($a.Value) { ": $($a.Value)" })" -ForegroundColor Cyan }
    }
}

# Find one rule by exact name (case-insensitive) or execution order number.
function Find-Rule {
    param($Rules, [string]$Name = "", [int]$Order = 0)
    for ($i = 1; $i -le $Rules.Count; $i++) {
        $r = $Rules.Item($i)
        if (($Name -and $r.Name -ieq $Name) -or ($Order -gt 0 -and [int]$r.ExecutionOrder -eq $Order)) { return $r }
    }
    $what = if ($Name) { "named '$Name'" } else { "at position $Order" }
    throw "No rule $what. List rules with outlook-rules-list.ps1."
}

# ---------------------------------------------------------------------------
# Composing: Outlook-native HTML, reply/forward insertion, mail builder
# ---------------------------------------------------------------------------

$script:NativeFont = "Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif"

# Plain text -> the HTML Outlook itself writes for new messages: one <div> per line,
# Aptos 11pt, black, class "elementToProof", and <div><br></div> for blank lines.
function ConvertTo-OutlookHtml {
    param([string]$Text, [string]$Font = $script:NativeFont, [string]$Size = "11pt")
    $style = "font-family: $Font; font-size: $Size; color: rgb(0, 0, 0);"
    $lines = ($Text -replace "`r", "").Split("`n")
    $sb = New-Object System.Text.StringBuilder
    foreach ($line in $lines) {
        if ($line.Trim() -eq "") { [void]$sb.Append("<div style=""$style"" class=""elementToProof""><br></div>") }
        else { [void]$sb.Append("<div style=""$style"" class=""elementToProof"">$([System.Net.WebUtility]::HtmlEncode($line))</div>") }
    }
    return $sb.ToString()
}

# Resolve the body a caller asked for into HTML (or $null for plain text).
#   -HtmlFile path | -HTML (Body is HTML) | -Native (Body is text, render Outlook-native) | text
function Get-ComposeBody {
    param([string]$Body, [switch]$HTML, [string]$HtmlFile, [switch]$Native)
    if ($HtmlFile) {
        if (-not (Test-Path -LiteralPath $HtmlFile)) { throw "HTML file not found: $HtmlFile" }
        return @{ Html = (Get-Content -LiteralPath $HtmlFile -Raw -Encoding UTF8); Text = $null }
    }
    if ($HTML) { return @{ Html = $Body; Text = $null } }
    if ($Native) { return @{ Html = (ConvertTo-OutlookHtml $Body); Text = $null } }
    return @{ Html = $null; Text = $Body }
}

# Remove the auto-signature Outlook inserts into replies/forwards (<a name="_MailAutoSig">),
# from its block up to the quoted header (<a name="_MailOriginal">). Returns the new HTML.
function Remove-AutoSignature {
    param([string]$Html)
    $sig = $Html.IndexOf("_MailAutoSig")
    if ($sig -lt 0) { return $Html }
    # Start of the run of opening tags that leads into the signature anchor.
    $start = $Html.LastIndexOf("<p", $sig)
    while ($start -gt 0) {
        $prev = $Html.LastIndexOf("<", $start - 1)
        if ($prev -ge 0 -and $Html.Substring($prev, $start - $prev) -match '^<div[^>]*>\s*$') { $start = $prev } else { break }
    }
    $orig = $Html.IndexOf("_MailOriginal")
    $end = if ($orig -gt $sig) { $Html.LastIndexOf("<div", $orig) } else { $Html.IndexOf("</body>", $sig) }
    if ($start -lt 0 -or $end -le $start) { return $Html }
    return $Html.Remove($start, $end - $start)
}

# Put the caller's HTML at the top of a reply/forward body (above signature and quoted thread),
# inside a marked block. If the block already exists (draft being updated), replace it.
function Set-ReplyBlock {
    param($Item, [string]$Html, [switch]$RemoveAutoSignature)
    $h = $Item.HTMLBody
    $block = "<!--suite-reply-start--><div id=""suite-reply"">$Html</div><!--suite-reply-end-->"
    $s = $h.IndexOf("<!--suite-reply-start-->"); $e = $h.IndexOf("<!--suite-reply-end-->")
    if ($s -ge 0 -and $e -gt $s) {
        $h = $h.Remove($s, $e + "<!--suite-reply-end-->".Length - $s).Insert($s, $block)
    } else {
        if ($RemoveAutoSignature) { $h = Remove-AutoSignature $h }
        $m = [regex]::Match($h, '<div class=WordSection1>')
        if (-not $m.Success) { $m = [regex]::Match($h, '<body[^>]*>') }
        $at = if ($m.Success) { $m.Index + $m.Length } else { 0 }
        $h = $h.Insert($at, $block)
    }
    $Item.HTMLBody = $h
}

# Apply compose options to a MailItem (new, reply, forward or existing draft).
function Set-MailFields {
    param(
        $Mail,
        [string]$To, [string]$CC, [string]$BCC, [string]$Subject,
        [string]$Account, [string]$Importance, [string]$Sensitivity,
        [switch]$ReadReceipt, [switch]$DeliveryReceipt, [string]$VotingOptions, $SendAt,
        [string[]]$Attachment = @()
    )
    if ($PSBoundParameters.ContainsKey("To") -and $To) { $Mail.To = $To }
    if ($PSBoundParameters.ContainsKey("CC") -and $CC) { $Mail.CC = $CC }
    if ($PSBoundParameters.ContainsKey("BCC") -and $BCC) { $Mail.BCC = $BCC }
    if ($Subject) { $Mail.Subject = $Subject }
    if ($Account) {
        $acct = $null
        foreach ($a in $Namespace.Accounts) {
            if ($a.SmtpAddress -ieq $Account -or $a.DisplayName -ieq $Account) { $acct = $a; break }
        }
        if (-not $acct) { throw "Account '$Account' not found. See outlook-accounts-list.ps1." }
        $Mail.SendUsingAccount = $acct
    }
    if ($Importance) { $Mail.Importance = @{ "Low" = 0; "Normal" = 1; "High" = 2 }[$Importance] }
    if ($Sensitivity) { $Mail.Sensitivity = @{ "Normal" = 0; "Personal" = 1; "Private" = 2; "Confidential" = 3 }[$Sensitivity] }
    if ($ReadReceipt) { $Mail.ReadReceiptRequested = $true }
    if ($DeliveryReceipt) { $Mail.OriginatorDeliveryReportRequested = $true }
    if ($VotingOptions) { $Mail.VotingOptions = $VotingOptions }
    if ($SendAt) { $Mail.DeferredDeliveryTime = [datetime]$SendAt }
    foreach ($f in $Attachment) {
        if (-not $f) { continue }
        if (-not (Test-Path -LiteralPath $f)) { throw "Attachment not found: $f" }
        [void]$Mail.Attachments.Add((Resolve-Path -LiteralPath $f).Path)
    }
}

# Create a new message from a script's bound parameters, then save it as a draft or send it.
# Sending is gated: without -Confirm (and with confirmation on) it previews and discards.
function Invoke-NewMail {
    param([hashtable]$P, [string]$Action = "draft", [switch]$Send, [string]$Next = "")
    $mail = $Outlook.CreateItem(0)
    $body = Get-ComposeBody -Body $P.Body -HTML:([bool]$P.HTML) -HtmlFile $P.HtmlFile -Native:([bool]$P.Native)
    if ($body.Html) { $mail.HTMLBody = $body.Html } elseif ($null -ne $body.Text) { $mail.Body = $body.Text }
    $fields = @{}
    foreach ($k in "To", "CC", "BCC", "Subject", "Account", "Importance", "Sensitivity", "VotingOptions", "SendAt", "Attachment", "ReadReceipt", "DeliveryReceipt") {
        if ($P.ContainsKey($k) -and $null -ne $P[$k] -and "$($P[$k])" -ne "") { $fields[$k] = $P[$k] }
    }
    if ($fields.ContainsKey("Attachment")) { $fields.Attachment = @($fields.Attachment) }
    Set-MailFields -Mail $mail @fields
    Assert-OutlookRecipients $mail

    if ($Send) {
        if (-not ([bool]$P.Confirm -or -not (Test-ConfirmRequired -Action $Action))) {
            Write-ComposeSummary -Mail $mail -Title "PREVIEW (not sent)" -Json:([bool]$P.Json) -Action $Action
            if (-not $P.Json) { Approve-Action -Action $Action -Confirmed $false -What "sending" | Out-Null }
            try { $mail.Close(1) } catch { }   # olDiscard
            return
        }
        $summaryMail = $mail
        $mail.Send()
        Write-ComposeSummary -Mail $summaryMail -Title "SENT" -Json:([bool]$P.Json) -Action $Action -Next "Find it in Sent Items: outlook-find.ps1 -Folder Sent -Days 1"
        return
    }
    $mail.Save()
    if ($P.Display) { $mail.Display() }
    Write-ComposeSummary -Mail $mail -Title "DRAFT SAVED" -Json:([bool]$P.Json) -Action $Action -Next $(if ($Next) { $Next } else { "Review in Drafts; send with outlook-send-draft.ps1 -EntryID <EntryID> -Confirm" })
}

# Text summary of a composed message (draft or about to send).
function Write-ComposeSummary {
    param($Mail, [string]$Title, [string]$Original = "", [switch]$Json, [string]$Action = "", [string]$Next = "")
    $att = @(); for ($i = 1; $i -le $Mail.Attachments.Count; $i++) { $att += $Mail.Attachments.Item($i).FileName }
    $preview = Format-Preview $Mail.Body 300
    $o = [ordered]@{
        Result = $Title; Action = $Action; EntryID = $Mail.EntryID; Subject = $Mail.Subject
        To = $Mail.To; CC = $Mail.CC; BCC = $Mail.BCC
        From = $(try { $Mail.SendUsingAccount.SmtpAddress } catch { "" })
        Format = @{ 1 = "text"; 2 = "html"; 3 = "rtf" }[[int]$Mail.BodyFormat]
        Attachments = $att; Preview = $preview
    }
    if ($Mail.DeferredDeliveryTime -and ([datetime]$Mail.DeferredDeliveryTime).Year -lt 4500) { $o.SendAt = Format-MailDateIso $Mail.DeferredDeliveryTime }
    if ($Original) { $o.InReplyTo = $Original }
    if ($Json) { Write-Json $o; return }
    Write-Host "=== $Title ===" -ForegroundColor Green
    if ($o.EntryID) { Write-Host "  EntryID: $($o.EntryID)" }
    if ($Original) { Write-Host "  Original: $Original" -ForegroundColor DarkGray }
    Write-Host "  Subject: $($o.Subject)" -ForegroundColor Yellow
    if ($o.From) { Write-Host "  From account: $($o.From)" }
    Write-Host "  To: $($o.To)"
    if ($o.CC) { Write-Host "  CC: $($o.CC)" }
    if ($o.BCC) { Write-Host "  BCC: $($o.BCC)" }
    Write-Host "  Format: $($o.Format)"
    if ($o.SendAt) { Write-Host "  Send at: $(Format-MailDate $Mail.DeferredDeliveryTime)" -ForegroundColor Cyan }
    if ($att.Count) { Write-Host "  Attachments ($($att.Count)): $($att -join ', ')" }
    if ($preview) { Write-Host "  Preview: $preview" -ForegroundColor Gray }
    if ($Next) { Write-Host "  Next: $Next" -ForegroundColor DarkGray }
}

# ---------------------------------------------------------------------------
# Confirmation policy for irreversible or outward-facing actions
# ---------------------------------------------------------------------------

# Order of precedence:
#   1. env OUTLOOK_SUITE_REQUIRE_CONFIRM = 0 or 1
#   2. %APPDATA%\outlook-suite\config.json  { "requireConfirm": true, "actions": { "send": false } }
#   3. default: confirmation required
function Test-ConfirmRequired {
    param([Parameter(Mandatory = $true)][string]$Action)
    $envValue = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
    if ($envValue -in @("0", "false", "False")) { return $false }
    if ($envValue -in @("1", "true", "True")) { return $true }
    $cfgPath = Join-Path $env:APPDATA "outlook-suite\config.json"
    if (Test-Path $cfgPath) {
        try {
            $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
            if ($cfg.actions -and $null -ne $cfg.actions.$Action) { return [bool]$cfg.actions.$Action }
            if ($null -ne $cfg.requireConfirm) { return [bool]$cfg.requireConfirm }
        } catch { }
    }
    return $true
}

# Call before an irreversible action. Returns $true when the action may run.
# When confirmation is required and -Confirm was not given, prints the preview
# footer and returns $false (the script then exits 0 without acting).
function Approve-Action {
    param([string]$Action, [bool]$Confirmed, [string]$What)
    if ($Confirmed -or -not (Test-ConfirmRequired -Action $Action)) { return $true }
    Write-Host ""
    Write-Host "PREVIEW ONLY: $What was NOT done." -ForegroundColor Yellow
    Write-Host "To do it, run the same command with -Confirm. (Confirmation for '$Action' can be turned off in %APPDATA%\outlook-suite\config.json or with OUTLOOK_SUITE_REQUIRE_CONFIRM=0.)" -ForegroundColor Gray
    return $false
}
