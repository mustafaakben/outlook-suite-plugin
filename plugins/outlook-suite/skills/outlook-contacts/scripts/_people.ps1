# People helpers for the outlook-contacts scripts (and any skill that needs to turn a
# name into a person). Dot-source after _common.ps1:
#   . (Join-Path $PSScriptRoot "_common.ps1")
#   . (Join-Path $PSScriptRoot "_people.ps1")        # from another skill:
#   . (Join-Path $PSScriptRoot "..\..\outlook-contacts\scripts\_people.ps1")
#
# Main entry points:
#   Resolve-Person      -Namespace $ns -Query "jane" [-Limit 10] [-Sources Contacts,Directory,Recent] [-Days 180] [-Detail]
#   Resolve-PersonSmtp  -Namespace $ns -Query "jane doe"   # one SMTP address, or throws listing the candidates
#   Get-ContactRecords  -Namespace $ns [-AllStores] [-Search "..."]
#
# Sources:
#   Contacts   default Contacts folder of every mailbox (one GetTable per folder)
#   Directory  the Exchange Global Address List (your organization's directory). Display names are
#              matched against a local name index (%LOCALAPPDATA%\outlook-suite\directory-names.txt),
#              built on first use (about 15 s for 17k entries) and refreshed after 7 days.
#   Recent     senders in every received folder (Inbox and rule-filed folders) and Sent Items
#              recipients within -Days (one indexed Table query per folder)

$script:ContactColumns = @("EntryID", "FullName", "CompanyName", "JobTitle", "Department", "OfficeLocation",
    "Email1Address", "Email1AddressType", "Email2Address", "Email3Address",
    "BusinessTelephoneNumber", "MobileTelephoneNumber", "HomeTelephoneNumber", "Categories", "MessageClass",
    "http://schemas.microsoft.com/mapi/proptag/0x66700102",   # long-term EntryID (the table's EntryID column is short-term)
    "CreationTime", "Email2AddressType", "Email3AddressType", "LastModificationTime")

# Contact fields the scripts read and write: script/record field -> ContactItem property.
$script:ContactFields = [ordered]@{
    FullName = "FullName"; Email = "Email1Address"; Email2 = "Email2Address"; Email3 = "Email3Address"
    Phone = "BusinessTelephoneNumber"; Mobile = "MobileTelephoneNumber"; HomePhone = "HomeTelephoneNumber"
    Company = "CompanyName"; JobTitle = "JobTitle"; Department = "Department"; Office = "OfficeLocation"
    Categories = "Categories"; Notes = "Body"
}

# ---------------------------------------------------------------------------
# Contacts
# ---------------------------------------------------------------------------

# Default Contacts folder of the default mailbox, or of every mailbox with -AllStores.
function Get-ContactFolders {
    param($Namespace, [switch]$AllStores)
    $list = New-Object System.Collections.ArrayList
    if (-not $AllStores) { [void]$list.Add($Namespace.GetDefaultFolder(10)); return , $list }
    foreach ($s in (Get-SearchStores -Namespace $Namespace)) {
        try { [void]$list.Add($s.GetDefaultFolder(10)) } catch { }   # some stores have no Contacts folder
    }
    return , $list
}

# SMTP for a contact email that may be an Exchange (EX) address. An EX address that does not
# resolve returns "" (never the /O=... legacy DN, which must not be shown, compared or copied).
function Resolve-ContactSmtp {
    param($Namespace, [string]$Address, [string]$Type)
    if (-not $Address) { return "" }
    if ($Type -ne "EX" -and $Address -notmatch '^/O=') { return $Address }
    try {
        $r = $Namespace.CreateRecipient($Address)
        if ($r.Resolve()) { $u = $r.AddressEntry.GetExchangeUser(); if ($u) { return [string]$u.PrimarySmtpAddress } }
    } catch { }
    return ""
}

# All contacts (no distribution lists) as ordered records, read with one GetTable per folder.
# -Search keeps contacts whose name, emails, company, title or department contain the text.
function Get-ContactRecords {
    param($Namespace, [switch]$AllStores, [string]$Search = "", [switch]$NoResolve)
    # -NoResolve keeps Exchange (EX) addresses as they are (no server lookup per contact);
    # Email1AddressType then says "EX" and the caller resolves only the contacts it keeps.
    $list = New-Object System.Collections.Generic.List[object]
    $needle = $Search.ToLowerInvariant()
    foreach ($f in (Get-ContactFolders -Namespace $Namespace -AllStores:$AllStores)) {
        try {
            $table = $f.GetTable()
            $table.Columns.RemoveAll()
            foreach ($c in $script:ContactColumns) { [void]$table.Columns.Add($c) }
            $table.Sort("FullName", $false)
            $n = $table.GetRowCount()
            if ($n -eq 0) { continue }
            $rows = $table.GetArray($n)
        } catch { continue }   # a mailbox that cannot be read now must not fail the whole list
        for ($i = 0; $i -lt $n; $i++) {
            if ([string]$rows[$i, 14] -notlike "IPM.Contact*") { continue }
            $rec = [ordered]@{
                EntryID       = $(if ($rows[$i, 15] -is [byte[]]) { ConvertTo-HexId $rows[$i, 15] } else { [string]$rows[$i, 0] })
                StoreID       = $f.StoreID
                FullName      = [string]$rows[$i, 1]
                Company       = [string]$rows[$i, 2]
                JobTitle      = [string]$rows[$i, 3]
                Department    = [string]$rows[$i, 4]
                Office        = [string]$rows[$i, 5]
                Email         = $(if ($NoResolve) { [string]$rows[$i, 6] } else { Resolve-ContactSmtp $Namespace ([string]$rows[$i, 6]) ([string]$rows[$i, 7]) })
                EmailType     = [string]$rows[$i, 7]
                Email2        = $(if ($NoResolve) { [string]$rows[$i, 8] } else { Resolve-ContactSmtp $Namespace ([string]$rows[$i, 8]) ([string]$rows[$i, 17]) })
                Email3        = $(if ($NoResolve) { [string]$rows[$i, 9] } else { Resolve-ContactSmtp $Namespace ([string]$rows[$i, 9]) ([string]$rows[$i, 18]) })
                BusinessPhone = [string]$rows[$i, 10]
                MobilePhone   = [string]$rows[$i, 11]
                HomePhone     = [string]$rows[$i, 12]
                Categories    = [string]$rows[$i, 13]
                Created       = (Format-MailDateIso $rows[$i, 16])   # ISO string: same JSON in 5.1 and 7, sorts by time
                Modified      = (Format-MailDateIso $rows[$i, 19])
                Folder        = $f.FolderPath
            }
            if ($needle) {
                $hay = (@($rec.FullName, $rec.Email, $rec.Email2, $rec.Email3, $rec.Company, $rec.JobTitle, $rec.Department) -join "`n").ToLowerInvariant()
                if (-not $hay.Contains($needle)) { continue }
            }
            $list.Add($rec)
        }
    }
    return , $list
}

# Contact item for the action scripts: -EntryID (preferred) or -Index, the 1-based position in
# outlook-contacts-list.ps1 (same -Search, same order). Throws with a hint when not found.
function Get-ContactTarget {
    param($Namespace, [string]$EntryID = "", [int]$Index = 0, [string]$Search = "")
    if ($EntryID) {
        $item = $null
        try { $item = $Namespace.GetItemFromID($EntryID) } catch { }
        if (-not $item) {
            foreach ($s in (Get-SearchStores -Namespace $Namespace)) {
                try { $item = $Namespace.GetItemFromID($EntryID, $s.StoreID); break } catch { }
            }
        }
        if (-not $item) { throw "No contact found with that EntryID. Run outlook-contacts-list.ps1 to get a current EntryID." }
        if ([string]$item.MessageClass -notlike "IPM.Contact*") { throw "That EntryID is not a contact (it is $($item.MessageClass))." }
        return $item
    }
    if ($Index -gt 0) {
        Write-Host "Note: -Index is the position in outlook-contacts-list.ps1 and shifts as contacts change. Prefer -EntryID." -ForegroundColor DarkGray
        $recs = Get-ContactRecords -Namespace $Namespace -Search $Search
        if ($Index -gt $recs.Count) { throw "No contact at -Index $Index ($($recs.Count) contacts$(if ($Search) { " matching '$Search'" }))." }
        return $Namespace.GetItemFromID($recs[$Index - 1].EntryID, $recs[$Index - 1].StoreID)
    }
    throw "Provide -EntryID (preferred) or -Index to choose a contact."
}

# Ordered record from a live ContactItem (same keys as Get-ContactRecords).
function ConvertTo-ContactRecord {
    param($Contact)
    $folder = ""; try { $folder = $Contact.Parent.FolderPath } catch { }
    return [ordered]@{
        EntryID = [string]$Contact.EntryID; StoreID = ""; FullName = [string]$Contact.FullName
        Company = [string]$Contact.CompanyName; JobTitle = [string]$Contact.JobTitle; Department = [string]$Contact.Department
        Office = [string]$Contact.OfficeLocation; Email = [string]$Contact.Email1Address; Email2 = [string]$Contact.Email2Address
        Email3 = [string]$Contact.Email3Address; BusinessPhone = [string]$Contact.BusinessTelephoneNumber
        MobilePhone = [string]$Contact.MobileTelephoneNumber; HomePhone = [string]$Contact.HomeTelephoneNumber
        Categories = [string]$Contact.Categories; Folder = $folder
    }
}

# Text view of one contact record. -Number prefixes the list position.
function Write-ContactRecord {
    param($Rec, [int]$Number = 0)
    $head = if ($Number -gt 0) { "[$Number] " } else { "" }
    $name = if ($Rec.FullName) { $Rec.FullName } else { "(no name)" }
    Write-Host "$head$name" -ForegroundColor Yellow
    $role = @($Rec.JobTitle, $Rec.Department, $Rec.Company) | Where-Object { $_ }
    if ($role) { Write-Host "    $($role -join ' | ')" -ForegroundColor Gray }
    foreach ($e in @($Rec.Email, $Rec.Email2, $Rec.Email3)) { if ($e) { Write-Host "    Email: $e" } }
    if ($Rec.BusinessPhone) { Write-Host "    Work: $($Rec.BusinessPhone)" }
    if ($Rec.MobilePhone) { Write-Host "    Mobile: $($Rec.MobilePhone)" }
    if ($Rec.HomePhone) { Write-Host "    Home: $($Rec.HomePhone)" }
    if ($Rec.Office) { Write-Host "    Office: $($Rec.Office)" }
    if ($Rec.Categories) { Write-Host "    Categories: $($Rec.Categories)" -ForegroundColor DarkGray }
    Write-Host "    EntryID: $($Rec.EntryID)" -ForegroundColor DarkGray
}

# Confirm gate for the contact scripts (delete, dedupe -Merge, import). Returns exactly one [bool]:
# $true when the action may run. When confirmation is required and missing it returns $false;
# in text mode it prints the PREVIEW ONLY footer (Write-Host, never the pipeline). In -Json mode
# it prints nothing and the caller writes the preview JSON. Callers: if (-not (Approve-ContactAction ...)) { preview; exit 0 }
function Approve-ContactAction {
    param([string]$Action, [bool]$Confirmed, [string]$What, [switch]$Json)
    if ($Confirmed -or -not (Test-ConfirmRequired -Action $Action)) { return $true }
    if (-not $Json) { $null = Approve-Action -Action $Action -Confirmed $false -What $What }
    return $false
}

# Result of a contact action (create, update, photo, export, delete).
function Write-ContactResult {
    param([string]$Title, $Rec, $Changes = [ordered]@{}, [string]$Action, [switch]$Json)
    if ($Json) {
        $o = [ordered]@{ Action = $Action; Result = $Title }
        foreach ($k in $Rec.Keys) { $o[$k] = $Rec[$k] }
        foreach ($k in $Changes.Keys) { $o[$k] = $Changes[$k] }
        Write-Json $o
        return
    }
    Write-Host "=== $Title ===" -ForegroundColor Green
    Write-ContactRecord $Rec
    foreach ($k in $Changes.Keys) { Write-Host "  ${k}: $($Changes[$k])" -ForegroundColor Cyan }
}

# ---------------------------------------------------------------------------
# Directory (Exchange Global Address List)
# ---------------------------------------------------------------------------

function Get-PeopleCache {
    if ($global:OutlookSuite) {
        if (-not $global:OutlookSuite.People) { $global:OutlookSuite.People = @{} }
        return $global:OutlookSuite.People
    }
    if (-not $script:PeopleCacheLocal) { $script:PeopleCacheLocal = @{} }
    return $script:PeopleCacheLocal
}

# The Exchange GAL with the most entries (a profile with several accounts lists one GAL each;
# GetGlobalAddressList() can return a small one from another account). $null when there is none.
function Get-DirectoryList {
    param($Namespace)
    $best = $null; $max = -1
    foreach ($al in $Namespace.AddressLists) {
        try {
            if ($al.AddressListType -ne 0) { continue }   # olExchangeGlobalAddressList
            $c = $al.AddressEntries.Count
            if ($c -gt $max) { $max = $c; $best = $al }
        } catch { }
    }
    return $best
}

# Display names of every directory entry. Read from the local index file when it is fresh,
# otherwise rebuilt by walking the GAL once. Kept in memory inside the warm host.
function Get-DirectoryNames {
    param($Namespace, [int]$MaxAgeDays = 7, [switch]$Refresh)
    $cache = Get-PeopleCache
    if (-not $Refresh -and $cache.DirectoryNames -and $cache.DirectoryLoaded -gt (Get-Date).AddHours(-12)) { return , $cache.DirectoryNames }
    $dir = Join-Path $env:LOCALAPPDATA "outlook-suite"
    $file = Join-Path $dir "directory-names.txt"
    $names = $null
    if (-not $Refresh -and (Test-Path -LiteralPath $file) -and (Get-Item -LiteralPath $file).LastWriteTime -gt (Get-Date).AddDays(-$MaxAgeDays)) {
        $names = [IO.File]::ReadAllLines($file)
    } else {
        $gal = Get-DirectoryList -Namespace $Namespace
        if (-not $gal) { return , @() }
        $list = New-Object System.Collections.Generic.List[string]
        $entries = $gal.AddressEntries
        $e = $entries.GetFirst()
        while ($e) { $list.Add([string]$e.Name); $e = $entries.GetNext() }
        $names = $list.ToArray()
        # Save only a complete walk (a dead session or a cut-short walk must not be trusted for
        # 7 days), and write to a temp file first so a concurrent run never reads half a file.
        $expected = 0; try { $expected = [int]$entries.Count } catch { }
        if ($names.Count -gt 0 -and $names.Count -ge 0.95 * $expected) {
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $tmp = "$file.$PID.tmp"
            [IO.File]::WriteAllLines($tmp, $names, (New-Object System.Text.UTF8Encoding($false)))
            Move-Item -LiteralPath $tmp -Destination $file -Force
        } else {
            return , $names   # use it for this call only; do not cache an incomplete index
        }
    }
    $cache.DirectoryNames = $names
    $cache.DirectoryLoaded = Get-Date
    return , $names
}

# Directory attribute index: one "Name<TAB>Department<TAB>Title" line per entry, for searching
# by department or title. Separate from the name index because it costs more to build (one
# batched property read per entry, about 20 s for 17k entries); built on first attribute search,
# refreshed after 7 days, saved only when complete, written atomically. Display data only.
function Get-DirectoryAttributes {
    param($Namespace, [int]$MaxAgeDays = 7, [switch]$Refresh)
    $cache = Get-PeopleCache
    if (-not $Refresh -and $cache.DirectoryAttrs -and $cache.AttrsLoaded -gt (Get-Date).AddHours(-12)) { return , $cache.DirectoryAttrs }
    $dir = Join-Path $env:LOCALAPPDATA "outlook-suite"
    $file = Join-Path $dir "directory-attributes.tsv"
    $lines = $null
    if (-not $Refresh -and (Test-Path -LiteralPath $file) -and (Get-Item -LiteralPath $file).LastWriteTime -gt (Get-Date).AddDays(-$MaxAgeDays)) {
        $lines = [IO.File]::ReadAllLines($file)
    } else {
        $gal = Get-DirectoryList -Namespace $Namespace
        if (-not $gal) { return , @() }
        $props = @("http://schemas.microsoft.com/mapi/proptag/0x3001001F",   # PR_DISPLAY_NAME
            "http://schemas.microsoft.com/mapi/proptag/0x3A18001F",          # PR_DEPARTMENT_NAME
            "http://schemas.microsoft.com/mapi/proptag/0x3A17001F")          # PR_TITLE
        $list = New-Object System.Collections.Generic.List[string]
        $entries = $gal.AddressEntries
        $e = $entries.GetFirst()
        while ($e) {
            $v = $null; try { $v = $e.PropertyAccessor.GetProperties($props) } catch { }
            $f = @("", "", "")
            if ($v) { for ($i = 0; $i -lt 3; $i++) { if ($v[$i] -is [string]) { $f[$i] = $v[$i] -replace '[\t\r\n]', ' ' } } }   # missing values come back as Int32 error codes
            if (-not $f[0]) { $f[0] = ([string]$e.Name) -replace '[\t\r\n]', ' ' }
            $list.Add($f -join "`t")
            $e = $entries.GetNext()
        }
        $lines = $list.ToArray()
        $expected = 0; try { $expected = [int]$entries.Count } catch { }
        if ($lines.Count -gt 0 -and $lines.Count -ge 0.95 * $expected) {
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $tmp = "$file.$PID.tmp"
            [IO.File]::WriteAllLines($tmp, $lines, (New-Object System.Text.UTF8Encoding($false)))
            Move-Item -LiteralPath $tmp -Destination $file -Force
        } else {
            return , $lines   # incomplete walk: this call only
        }
    }
    $cache.DirectoryAttrs = $lines
    $cache.AttrsLoaded = Get-Date
    return , $lines
}

# Directory entries whose department and/or title contain the given text (both given: both
# must match), optionally also matching a name query. Returns Name, Department, Title records,
# sorted by name, at most $Limit, plus the total.
function Find-DirectoryByAttribute {
    param($Namespace, [string]$Department = "", [string]$Title = "", [string]$Query = "", [int]$Limit = 25)
    $lines = Get-DirectoryAttributes -Namespace $Namespace
    $cand = $lines
    foreach ($t in @($Department, $Title) + @($Query -split '\s+')) {   # native prefilter on the whole line
        if ($t) { $cand = @($cand -like ("*" + [Management.Automation.WildcardPattern]::Escape($t) + "*")) }
    }
    $d = $Department.ToLowerInvariant(); $ti = $Title.ToLowerInvariant()
    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($l in $cand) {
        $f = $l.Split("`t")
        if ($f.Count -lt 3) { continue }
        if ($d -and -not $f[1].ToLowerInvariant().Contains($d)) { continue }
        if ($ti -and -not $f[2].ToLowerInvariant().Contains($ti)) { continue }
        if ($Query -and (Get-PersonMatchScore -Query $Query -Name $f[0] -Email "") -eq 0) { continue }
        $hits.Add([ordered]@{ Name = $f[0]; Department = $f[1]; Title = $f[2] })
    }
    $sorted = @($hits | Sort-Object { $_.Name })
    return [ordered]@{ Total = $sorted.Count; Records = @($sorted | Select-Object -First $Limit) }
}

# ---------------------------------------------------------------------------
# Org chart
# ---------------------------------------------------------------------------

function ConvertTo-OrgPerson {
    param($User)
    return [ordered]@{ Name = [string]$User.Name; Email = [string]$User.PrimarySmtpAddress; Title = [string]$User.JobTitle
        Department = [string]$User.Department; Office = [string]$User.OfficeLocation; BusinessPhone = [string]$User.BusinessTelephoneNumber }
}

# People from an AddressEntries collection (direct reports), at most $Limit, skipping $ExceptName.
# Only the people kept are looked up (one directory read each); the rest are only counted.
function ConvertTo-OrgList {
    param($Entries, [int]$Limit, [string]$ExceptName = "")
    $list = New-Object System.Collections.Generic.List[object]
    $total = 0
    if (-not $Entries) { return [ordered]@{ Total = 0; People = $list } }
    # Every entry read is a server round trip, so read only until $Limit people are kept; the
    # total comes from Count (minus the person, who is always among their manager's reports).
    $n = $Entries.Count
    $total = $n; if ($ExceptName) { $total = [Math]::Max(0, $n - 1) }
    for ($i = 1; $i -le $n -and $list.Count -lt $Limit; $i++) {
        $e = $Entries.Item($i)
        if ($ExceptName -and [string]$e.Name -eq $ExceptName) { continue }
        $u = $null; try { $u = $e.GetExchangeUser() } catch { }
        $rec = if ($u) { ConvertTo-OrgPerson $u } else { [ordered]@{ Name = [string]$e.Name; Email = ""; Title = ""; Department = ""; Office = ""; BusinessPhone = "" } }
        $list.Add($rec)
    }
    return [ordered]@{ Total = $total; People = $list }
}

# Org view for one directory person (ExchangeUser): manager chain (up to $Levels), direct reports
# and peers (the manager's other direct reports), each capped at $Limit. Read-only.
function Get-OrgChart {
    param($User, [int]$Levels = 4, [int]$Limit = 25, [int]$PeerLimit = 10)
    $chain = New-Object System.Collections.Generic.List[object]
    $m = $null; try { $m = $User.GetExchangeUserManager() } catch { }
    $manager = $m
    # Each level is one server lookup (about 0.5 s); none is made past -Levels.
    while ($m -and $chain.Count -lt $Levels) {
        $chain.Add((ConvertTo-OrgPerson $m))
        if ($chain.Count -ge $Levels) { break }
        $next = $null; try { $next = $m.GetExchangeUserManager() } catch { }
        if ($next -and $next.PrimarySmtpAddress -eq $m.PrimarySmtpAddress) { break }   # top of the tree points at itself
        $m = $next
    }
    $reports = $null; try { $reports = $User.GetDirectReports() } catch { }
    $peers = $null; if ($manager) { try { $peers = $manager.GetDirectReports() } catch { } }
    return [ordered]@{
        Person        = (ConvertTo-OrgPerson $User)
        Managers      = $chain
        DirectReports = (ConvertTo-OrgList -Entries $reports -Limit $Limit)
        Peers         = (ConvertTo-OrgList -Entries $peers -Limit $PeerLimit -ExceptName ([string]$User.Name))
    }
}

# True for Exchange directory entries: user (0), distribution list (1), remote user (5).
# CreateRecipient(...).Resolve() also succeeds for any one-off SMTP address; those are not directory hits.
function Test-ExchangeEntry {
    param($Entry)
    if (-not $Entry) { return $false }
    $t = -1; try { $t = [int]$Entry.AddressEntryUserType } catch { }
    return ($t -in 0, 1, 5)
}

# Details for one directory entry. -Detail adds the manager (one more server lookup).
function ConvertTo-DirectoryPerson {
    param($Entry, [switch]$Detail)
    $rec = [ordered]@{ Name = [string]$Entry.Name; Email = ""; Kind = "Person"; Title = ""; Department = ""; Company = ""
        Office = ""; BusinessPhone = ""; MobilePhone = ""; Manager = ""; Alias = "" }
    $type = -1; try { $type = [int]$Entry.AddressEntryUserType } catch { }
    if ($type -eq 1) {   # olExchangeDistributionListAddressEntry
        $rec.Kind = "Group"
        try { $dl = $Entry.GetExchangeDistributionList(); if ($dl) { $rec.Email = [string]$dl.PrimarySmtpAddress; $rec.Alias = [string]$dl.Alias } } catch { }
        return $rec
    }
    try {
        $u = $Entry.GetExchangeUser()
        if ($u) {
            $rec.Email = [string]$u.PrimarySmtpAddress; $rec.Title = [string]$u.JobTitle
            $rec.Department = [string]$u.Department; $rec.Company = [string]$u.CompanyName
            $rec.Office = [string]$u.OfficeLocation; $rec.BusinessPhone = [string]$u.BusinessTelephoneNumber
            $rec.MobilePhone = [string]$u.MobileTelephoneNumber; $rec.Alias = [string]$u.Alias
            if ($Detail) { try { $m = $u.GetExchangeUserManager(); if ($m) { $rec.Manager = [string]$m.Name } } catch { } }
        }
    } catch { }
    if ($type -in 2, 3) { $rec.Kind = "Public folder" }
    if (-not $rec.Email) { try { $rec.Email = [string]$Entry.Address } catch { } }
    return $rec
}

# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------

# How well a name/email matches the query: 100 exact, 80 every query word starts a name word
# (or the email local part), 60 substring, 0 no match.
function Get-PersonMatchScore {
    param([string]$Query, [string]$Name, [string]$Email)
    $q = $Query.Trim().ToLowerInvariant()
    if (-not $q) { return 0 }
    $n = ([string]$Name).ToLowerInvariant(); $e = ([string]$Email).ToLowerInvariant()
    if ($q -eq $n -or ($e -and $q -eq $e)) { return 100 }
    $words = @($n -split '[\s,.\-''()]+' | Where-Object { $_ })
    if ($e) { $words += @(($e.Split('@')[0]) -split '[.\-_]+' | Where-Object { $_ }) }
    $all = $true
    foreach ($t in ($q -split '\s+')) {
        $hit = $false
        foreach ($w in $words) { if ($w.StartsWith($t)) { $hit = $true; break } }
        if (-not $hit) { $all = $false; break }
    }
    if ($all) { return 80 }
    if ($n.Contains($q) -or ($e -and $e.Contains($q))) { return 60 }
    return 0
}

# Directory names matching the query (word-prefix or substring), best first, at most $Limit.
# Each hit: Name, Score, Count (how many directory entries share that display name).
function Find-DirectoryNames {
    param([string[]]$Names, [string]$Query, [int]$Limit = 10)
    $hits = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    # Native prefilter (array -like runs in C#): keep names containing every query word.
    $candidates = $Names
    foreach ($t in ($Query.Trim() -split '\s+')) {
        $candidates = @($candidates -like ("*" + [Management.Automation.WildcardPattern]::Escape($t) + "*"))
    }
    foreach ($n in $candidates) {
        # Namesakes: the index holds one line per entry, so equal display names are counted.
        # AddressEntries.Item(name) can open only the first, so callers treat Count > 1 as ambiguous.
        if ($seen.ContainsKey($n)) { $seen[$n].Count++; continue }
        $s = Get-PersonMatchScore -Query $Query -Name $n -Email ""
        if ($s -gt 0) { $h = [pscustomobject]@{ Name = $n; Score = $s; Count = 1 }; $seen[$n] = $h; $hits.Add($h) }
    }
    return , @($hits | Sort-Object @{ Expression = "Score"; Descending = $true }, Name | Select-Object -First $Limit)
}

# ---------------------------------------------------------------------------
# Received folders and recent correspondents
# ---------------------------------------------------------------------------

# Mail folders that hold received mail: the Inbox and the folders rules file mail into, in the
# default mailbox. Skips Sent, Drafts, Deleted, Junk, Outbox, sync issues, and non-mail
# folders, pruning whole subtrees (so it walks only what it returns). Cached for 10 minutes.
function Get-ReceivedFolders {
    param($Namespace)
    $cache = Get-PeopleCache
    if ($cache.ReceivedFolders -and $cache.ReceivedLoaded -gt (Get-Date).AddMinutes(-10)) { return , $cache.ReceivedFolders }
    $skip = @{}
    foreach ($id in 3, 4, 5, 16, 20, 23) { try { $skip[$Namespace.GetDefaultFolder($id).EntryID] = 1 } catch { } }
    $skipNames = @("Conversation History", "Yammer Root", "RSS Feeds", "Sync Issues")
    $out = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Stack
    foreach ($f in $Namespace.DefaultStore.GetRootFolder().Folders) { $stack.Push($f) }
    while ($stack.Count -gt 0) {
        $f = $stack.Pop()
        $type = -1; try { $type = [int]$f.DefaultItemType } catch { }
        if ($type -ne 0) { continue }                     # olMailItem folders only
        if ($skip.ContainsKey($f.EntryID) -or $skipNames -contains $f.Name) { continue }
        $out.Add($f)
        foreach ($sub in $f.Folders) { $stack.Push($sub) }
    }
    $cache.ReceivedFolders = $out
    $cache.ReceivedLoaded = Get-Date
    return , $out
}

# Received mail matching a DASL condition in every received folder: one indexed Table query
# per folder. Returns rows as [SenderName, SenderSmtp, ReceivedTime, Subject].
function Find-ReceivedRows {
    param($Namespace, [string]$Condition, [int]$Days)
    $since = (Get-Date).AddDays(-$Days).ToUniversalTime().ToString("yyyy-MM-dd HH:mm")
    $filter = "@SQL=($Condition) AND ""urn:schemas:httpmail:datereceived"" >= '$since'"
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($f in (Get-ReceivedFolders -Namespace $Namespace)) {
        $t = $f.GetTable($filter)
        $t.Columns.RemoveAll()
        foreach ($c in @("SenderName", $script:Prop.SenderSmtp, "ReceivedTime", "Subject")) { [void]$t.Columns.Add($c) }
        $n = $t.GetRowCount()
        if ($n -eq 0) { continue }
        $rows = $t.GetArray($n)
        for ($i = 0; $i -lt $n; $i++) {
            $r = @([string]$rows[$i, 0], [string]$rows[$i, 1], $rows[$i, 2], [string]$rows[$i, 3])
            $list.Add($r)
        }
    }
    return , $list
}

# People the query matches among senders of received mail (all received folders) and
# recipients of Sent Items, within -Days. Table queries only: Sent Items recipients come from
# the To/CC display lines, and names without an address are resolved through Exchange once each.
# Returns a hashtable keyed by lower-case SMTP (or "name:<name>"): Name, Email, Received, Sent, LastContact.
function Find-RecentCorrespondents {
    param($Namespace, [string]$Query, [int]$Days = 180)
    $out = @{}
    $q = ConvertTo-DaslText $Query.Trim()
    $byName = @{}   # lower-case name -> key

    $cond = """urn:schemas:httpmail:fromname"" LIKE '%$q%' OR ""urn:schemas:httpmail:fromemail"" LIKE '%$q%' OR ""$($script:Prop.SenderSmtp)"" LIKE '%$q%'"
    foreach ($r in (Find-ReceivedRows -Namespace $Namespace -Condition $cond -Days $Days)) {
        $name = $r[0]; $smtp = $r[1]
        if (-not $smtp) { continue }
        $k = $smtp.ToLowerInvariant()
        if (-not $out.ContainsKey($k)) { $out[$k] = [ordered]@{ Name = $name; Email = $smtp; Received = 0; Sent = 0; LastContact = $null } }
        $p = $out[$k]; $p.Received++
        if (-not $p.LastContact -or $r[2] -gt $p.LastContact) { $p.LastContact = $r[2] }
        if ($name) { $byName[$name.ToLowerInvariant()] = $k }
    }

    $since = (Get-Date).AddDays(-$Days).ToUniversalTime().ToString("yyyy-MM-dd HH:mm")
    $sent = $Namespace.GetDefaultFolder(5)
    $f = "@SQL=(""urn:schemas:httpmail:displayto"" LIKE '%$q%' OR ""urn:schemas:httpmail:displaycc"" LIKE '%$q%') AND ""urn:schemas:httpmail:date"" >= '$since'"
    $t = $sent.GetTable($f)
    $t.Columns.RemoveAll()
    foreach ($c in @("To", "http://schemas.microsoft.com/mapi/proptag/0x0E03001F", "SentOn")) { [void]$t.Columns.Add($c) }   # To, CC display lines
    $n = $t.GetRowCount()
    if ($n -gt 0) {
        $rows = $t.GetArray($n)
        for ($i = 0; $i -lt $n; $i++) {
            $d = $rows[$i, 2]
            $names = @(([string]$rows[$i, 0] + ";" + [string]$rows[$i, 1]) -split ';' | ForEach-Object { $_.Trim().Trim("'") } | Where-Object { $_ } | Select-Object -Unique)
            foreach ($nm in $names) {
                if ((Get-PersonMatchScore -Query $Query -Name $nm -Email $(if ($nm -match '@') { $nm } else { "" })) -eq 0) { continue }
                $ln = $nm.ToLowerInvariant()
                if (-not $byName.ContainsKey($ln)) {
                    $smtp = ""
                    if ($nm -match '^[^@\s]+@[^@\s]+$') { $smtp = $nm }
                    else {
                        try {
                            $rc = $Namespace.CreateRecipient($nm)
                            if ($rc.Resolve()) {
                                $ae = $rc.AddressEntry; $u = $null; try { $u = $ae.GetExchangeUser() } catch { }
                                $smtp = if ($u) { [string]$u.PrimarySmtpAddress } else { [string]$ae.Address }
                            }
                        } catch { }
                    }
                    $k = if ($smtp) { $smtp.ToLowerInvariant() } else { "name:$ln" }
                    $byName[$ln] = $k
                    if (-not $out.ContainsKey($k)) { $out[$k] = [ordered]@{ Name = $nm; Email = $smtp; Received = 0; Sent = 0; LastContact = $null } }
                }
                $p = $out[$byName[$ln]]; $p.Sent++
                if (-not $p.LastContact -or $d -gt $p.LastContact) { $p.LastContact = $d }
            }
        }
    }
    return $out
}

# ---------------------------------------------------------------------------
# Groups: Exchange distribution lists and local contact groups
# ---------------------------------------------------------------------------

# Member record of a directory entry or a contact-group recipient.
function ConvertTo-MemberRecord {
    param($Entry)
    $rec = [ordered]@{ Name = [string]$Entry.Name; Email = ""; Kind = "Person" }
    $type = -1; try { $type = [int]$Entry.AddressEntryUserType } catch { }
    if ($type -eq 1) {
        $rec.Kind = "Group"
        try { $dl = $Entry.GetExchangeDistributionList(); if ($dl) { $rec.Email = [string]$dl.PrimarySmtpAddress } } catch { }
    } else {
        try { $u = $Entry.GetExchangeUser(); if ($u) { $rec.Email = [string]$u.PrimarySmtpAddress } } catch { }
    }
    if (-not $rec.Email) { try { $a = [string]$Entry.Address; if ($a -notmatch '^/O=') { $rec.Email = $a } } catch { } }
    return $rec
}

# Members of a group given by name or address: a local contact group (IPM.DistList) in any
# Contacts folder first, then an Exchange distribution list. Read-only.
# Returns @{ Name; Email; Source; Total; Members } or throws when no group matches.
function Get-GroupMembers {
    param($Namespace, [Parameter(Mandatory = $true)][string]$Group, [int]$Limit = 100)
    $g = $Group.Trim(); $gl = $g.ToLowerInvariant()

    # Local contact groups
    foreach ($f in (Get-ContactFolders -Namespace $Namespace -AllStores)) {
        $t = $f.GetTable("[MessageClass] = 'IPM.DistList'")
        $t.Columns.RemoveAll(); [void]$t.Columns.Add("Subject"); [void]$t.Columns.Add($script:Prop.LongTermEntryID)
        while (-not $t.EndOfTable) {
            $row = $t.GetNextRow()
            if (([string]$row.Item(1)).ToLowerInvariant() -ne $gl) { continue }
            $id = $row.Item(2); if ($id -is [byte[]]) { $id = ConvertTo-HexId $id }
            $dl = $Namespace.GetItemFromID([string]$id, $f.StoreID)
            $members = New-Object System.Collections.Generic.List[object]
            $total = [int]$dl.MemberCount
            for ($i = 1; $i -le [Math]::Min($total, $Limit); $i++) {
                $r = $dl.GetMember($i)
                $rec = $null
                try { if ($r.AddressEntry) { $rec = ConvertTo-MemberRecord $r.AddressEntry } } catch { }
                if (-not $rec) { $rec = [ordered]@{ Name = [string]$r.Name; Email = [string]$r.Address; Kind = "Person" } }
                if (-not $rec.Name) { $rec.Name = [string]$r.Name }
                $members.Add($rec)
            }
            return [ordered]@{ Name = [string]$dl.DLName; Email = ""; Source = "Contact group"; Total = $total; Members = $members }
        }
    }

    # Exchange distribution lists: by address, by exact directory name, then by Exchange resolution.
    $entry = $null
    try { $r = $Namespace.CreateRecipient($g); if ($r.Resolve() -and (Test-ExchangeEntry $r.AddressEntry)) { $entry = $r.AddressEntry } } catch { }
    if (-not $entry -or [int]$entry.AddressEntryUserType -ne 1) {
        $gal = Get-DirectoryList -Namespace $Namespace
        if ($gal) {
            foreach ($h in (Find-DirectoryNames -Names (Get-DirectoryNames -Namespace $Namespace) -Query $g -Limit 10)) {
                try {
                    $e = $gal.AddressEntries.Item($h.Name)
                    if ($e -and $e.Name -eq $h.Name -and [int]$e.AddressEntryUserType -eq 1) { $entry = $e; break }
                } catch { }
            }
        }
    }
    if (-not $entry) { throw "No contact group or distribution list found for '$g'." }
    if ([int]$entry.AddressEntryUserType -ne 1) { throw "'$g' is a person, not a distribution list or contact group." }
    $dl = $entry.GetExchangeDistributionList()
    $entries = $dl.GetExchangeDistributionListMembers()
    $members = New-Object System.Collections.Generic.List[object]
    $total = 0
    if ($entries) {
        $total = $entries.Count
        for ($i = 1; $i -le [Math]::Min($total, $Limit); $i++) { $members.Add((ConvertTo-MemberRecord $entries.Item($i))) }
    }
    return [ordered]@{ Name = [string]$dl.Name; Email = [string]$dl.PrimarySmtpAddress; Source = "Directory"; Total = $total; Members = $members }
}
# ---------------------------------------------------------------------------
# Resolve-Person
# ---------------------------------------------------------------------------

# Find people matching a name, partial name or email across Contacts, the directory and
# recent mail. Returns a list of ordered records, best first:
#   Name, Email, Score (0-100+, ranking), Match (0-100, name/email match only),
#   Sources ("Contacts","Directory","Recent"), Kind, Title,
#   Department, Company, Office, BusinessPhone, MobilePhone, HomePhone, Manager,
#   OtherEmails, ContactEntryID, Received, Sent, LastContact,
#   Namesakes (directory entries sharing this display name; > 1 means the name alone is ambiguous)
function Resolve-Person {
    param(
        $Namespace,
        [Parameter(Mandatory = $true)][string]$Query,
        [int]$Limit = 10,
        [string[]]$Sources = @("Contacts", "Directory", "Recent"),
        [int]$Days = 180,
        [switch]$Detail
    )
    $Query = $Query.Trim()
    $people = [ordered]@{}   # key: lower-case email (or "name:<name>" when there is no email)

    $add = {
        param($rec, [string]$source, [int]$score)
        $key = if ($rec.Email) { ([string]$rec.Email).ToLowerInvariant() } else { "name:" + ([string]$rec.Name).ToLowerInvariant() }
        if (-not $people.Contains($key)) {
            $people[$key] = [ordered]@{ Name = [string]$rec.Name; Email = [string]$rec.Email; Score = 0; Match = 0; Sources = @()
                Kind = "Person"; Title = ""; Department = ""; Company = ""; Office = ""; BusinessPhone = ""; MobilePhone = ""
                HomePhone = ""; Manager = ""; OtherEmails = @(); ContactEntryID = ""; Received = 0; Sent = 0; LastContact = $null; Namesakes = 1 }
        }
        $p = $people[$key]
        if ($p.Sources -notcontains $source) { $p.Sources += $source }
        if ($score -gt $p.Score) { $p.Score = $score }
        if ($score -gt $p.Match) { $p.Match = $score }   # name/email match quality only, before ranking bonuses
        if ($rec.Contains("Namesakes") -and $rec.Namesakes -gt $p.Namesakes) { $p.Namesakes = $rec.Namesakes }
        foreach ($k in @("Kind", "Title", "Department", "Company", "Office", "BusinessPhone", "MobilePhone", "HomePhone", "Manager", "ContactEntryID", "Received", "Sent", "LastContact")) {
            if ($rec.Contains($k) -and $rec[$k] -and -not $p[$k]) { $p[$k] = $rec[$k] }
        }
        if ($source -eq "Directory" -and $rec.Name) { $p.Name = [string]$rec.Name }   # the directory spelling wins
    }

    $isEmail = $Query -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    if ($Sources -contains "Contacts") {
        foreach ($c in (Get-ContactRecords -Namespace $Namespace -AllStores -NoResolve)) {
            $s = Get-PersonMatchScore -Query $Query -Name $c.FullName -Email $(if ($c.EmailType -eq "EX") { "" } else { $c.Email })
            foreach ($e in @($c.Email2, $c.Email3)) { if ($e -and $e -notmatch '^/O=') { $s = [Math]::Max($s, (Get-PersonMatchScore -Query $Query -Name "" -Email $e)) } }
            if ($s -eq 0) {
                $hay = (@($c.Company, $c.JobTitle, $c.Department) -join " ").ToLowerInvariant()
                if ($hay.Contains($Query.ToLowerInvariant())) { $s = 40 } else { continue }
            }
            if ($c.EmailType -eq "EX") { $c.Email = Resolve-ContactSmtp $Namespace $c.Email "EX" }   # only for contacts that matched
            $rec = [ordered]@{ Name = $c.FullName; Email = $c.Email; Title = $c.JobTitle; Department = $c.Department; Company = $c.Company
                Office = $c.Office; BusinessPhone = $c.BusinessPhone; MobilePhone = $c.MobilePhone; HomePhone = $c.HomePhone; ContactEntryID = $c.EntryID }
            & $add $rec "Contacts" $s
            $key = if ($c.Email) { $c.Email.ToLowerInvariant() } else { "name:" + $c.FullName.ToLowerInvariant() }
            foreach ($e in @($c.Email2, $c.Email3)) { if ($e -and $e -notmatch '^/O=' -and $people[$key].OtherEmails -notcontains $e) { $people[$key].OtherEmails += $e } }   # never a legacy DN
        }
    }

    if ($Sources -contains "Directory") {
        $gal = Get-DirectoryList -Namespace $Namespace
        if ($gal) {
            $entries = $gal.AddressEntries
            $found = New-Object System.Collections.Generic.List[object]
            if ($isEmail) {
                try {
                    $r = $Namespace.CreateRecipient($Query)
                    if ($r.Resolve() -and (Test-ExchangeEntry $r.AddressEntry)) { $found.Add([pscustomobject]@{ Entry = $r.AddressEntry; Score = 100; Count = 1 }) }
                } catch { }
            } else {
                foreach ($h in (Find-DirectoryNames -Names (Get-DirectoryNames -Namespace $Namespace) -Query $Query -Limit $Limit)) {
                    try { $e = $entries.Item($h.Name); if ($e -and $e.Name -eq $h.Name) { $found.Add([pscustomobject]@{ Entry = $e; Score = $h.Score; Count = $h.Count }) } } catch { }
                }
                if ($found.Count -eq 0) {
                    # Not in the name index (alias, or the index is older than the entry): let Exchange resolve it.
                    try {
                        $r = $Namespace.CreateRecipient($Query)
                        if ($r.Resolve() -and (Test-ExchangeEntry $r.AddressEntry)) { $found.Add([pscustomobject]@{ Entry = $r.AddressEntry; Score = 70; Count = 1 }) }
                    } catch { }
                }
            }
            foreach ($h in $found) {
                $rec = ConvertTo-DirectoryPerson -Entry $h.Entry -Detail:$Detail
                $s = [Math]::Max($h.Score, (Get-PersonMatchScore -Query $Query -Name $rec.Name -Email $rec.Email))
                $rec.Namesakes = [int]$h.Count
                & $add $rec "Directory" $s
            }
        }
    }

    if ($Sources -contains "Recent" -and $Query.Length -ge 2) {
        $recent = Find-RecentCorrespondents -Namespace $Namespace -Query $Query -Days $Days
        foreach ($r in $recent.Values) {
            $s = Get-PersonMatchScore -Query $Query -Name $r.Name -Email $r.Email
            if ($s -eq 0) { $s = 50 }
            & $add $r "Recent" $s
            $key = if ($r.Email) { $r.Email.ToLowerInvariant() } else { "name:" + ([string]$r.Name).ToLowerInvariant() }
            $p = $people[$key]
            $p.Received = $r.Received; $p.Sent = $r.Sent; $p.LastContact = $r.LastContact
        }
    }

    # Rank: match score, +5 per extra source, + up to 10 for how often you correspond.
    $ranked = foreach ($p in $people.Values) {
        $p.Score = [int]($p.Score + 5 * ($p.Sources.Count - 1) + [Math]::Min(10, $p.Received + 2 * $p.Sent))
        $p
    }
    $sorted = @($ranked | Sort-Object @{ Expression = { $_.Score }; Descending = $true }, @{ Expression = { if ($_.LastContact) { $_.LastContact } else { [datetime]::MinValue } }; Descending = $true } | Select-Object -First $Limit)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($p in $sorted) { $list.Add($p) }
    return , $list
}

# Decides whether a ranked person list names one person clearly enough to act on (meeting
# attendees). Uses only Match (name/email match quality before ranking bonuses):
#   - the best Match is at least 80 (every query word starts a name word, or exact), and
#   - exactly one person has it, and that person has no namesake in the directory, and
#   - it is exact (100) or leads the next person's Match by 20 or more.
# Returns @{ Person = <record or $null>; Reason = "clear"|"ambiguous"|"namesakes"|"weak"|"none"; Candidates }.
function Select-ClearPerson {
    param($People)
    $hits = @($People | Where-Object { $_.Email })
    if ($hits.Count -eq 0) { return @{ Person = $null; Reason = "none"; Candidates = @() } }
    $byMatch = @($hits | Sort-Object { $_.Match } -Descending)
    $top = $byMatch[0]; $best = [int]$top.Match
    $second = if ($byMatch.Count -gt 1) { [int]$byMatch[1].Match } else { 0 }
    $cands = @($byMatch | Select-Object -First 5)
    if ($best -lt 80) { return @{ Person = $null; Reason = "weak"; Candidates = $cands } }
    if ($second -eq $best) { return @{ Person = $null; Reason = "ambiguous"; Candidates = $cands } }
    if ([int]$top.Namesakes -gt 1) { return @{ Person = $null; Reason = "namesakes"; Candidates = $cands } }
    if ($best -lt 100 -and $best - $second -lt 20) { return @{ Person = $null; Reason = "ambiguous"; Candidates = $cands } }
    return @{ Person = $top; Reason = "clear"; Candidates = $cands }
}

# One SMTP address for a name or email, for meeting attendees. An address is returned as is.
# Otherwise the person Select-ClearPerson accepts; it never guesses. Throws listing the
# candidates when the name is ambiguous (including namesakes), weak, or unknown.
function Resolve-PersonSmtp {
    param($Namespace, [Parameter(Mandatory = $true)][string]$Query, [int]$Days = 180)
    $q = $Query.Trim()
    if ($q -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return $q }
    $found = Resolve-Person -Namespace $Namespace -Query $q -Limit 10 -Days $Days   # assign first: the list comes back comma-wrapped
    $pick = Select-ClearPerson $found
    if ($pick.Person) { return [string]$pick.Person.Email }
    $names = ($pick.Candidates | ForEach-Object { "$($_.Name) <$($_.Email)>" }) -join "; "
    switch ($pick.Reason) {
        "none" { throw "No person found for '$q'. Give an email address instead." }
        "namesakes" { throw "'$q' matches more than one person: $($pick.Candidates[0].Namesakes) directory entries are named '$($pick.Candidates[0].Name)'. Use the email address." }
        "ambiguous" { throw "'$q' matches more than one person: $names. Use the email address." }
        default { throw "'$q' matches no one closely enough (best: $names). Use the email address." }
    }
}