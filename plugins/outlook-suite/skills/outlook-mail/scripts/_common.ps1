# Shared helpers for the outlook-mail scripts.
# Dot-source at the top of a script:  . (Join-Path $PSScriptRoot "_common.ps1")
#
# Works in two modes:
#   - standalone (pwsh -File script.ps1 ...): connects to Outlook once per call.
#   - inside outlook-host.ps1: reuses the warm Outlook session in $global:OutlookSuite.

Set-StrictMode -Off

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
    if ($Message -is [System.Management.Automation.ErrorRecord]) { $Message = $Message.Exception.Message }
    Write-Host "Error: $Message" -ForegroundColor Red
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
        try { return $Namespace.GetFolderFromID($cache[$key][0], $cache[$key][1]) } catch { $cache.Remove($key) }
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
        try {
            if ($StoreID) { return $Namespace.GetItemFromID($EntryID, $StoreID) }
            return $Namespace.GetItemFromID($EntryID)
        } catch {
            throw "No item found with EntryID $($EntryID.Substring(0, [Math]::Min(16, $EntryID.Length)))... It may have been moved (moving can change the EntryID) or deleted. Run outlook-find.ps1 again to get a current EntryID."
        }
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
    if ($Record.StoreID) { Write-Host "  StoreID: $($Record.StoreID)" -ForegroundColor DarkGray }
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
