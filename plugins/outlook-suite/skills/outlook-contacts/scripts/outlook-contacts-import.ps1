param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string]$Map = "",
    [switch]$Update,
    [switch]$Overwrite,
    [switch]$Confirm,
    [int]$Show = 20,
    [int]$MaxRows = 2000,
    [double]$MaxFileMB = 20,
    [switch]$Json
)

# Outlook Contacts Import: contacts from a vCard (.vcf, one or many cards) or a CSV file into
# the default Contacts folder. Previews by default (counts plus the first -Show rows, with the
# field changes of updates); writes only with -Confirm (confirmation key "import").
#
# Matching: every email of a row (Email, Email2, Email3) is looked up among your contacts.
#   - one contact matches: skipped, or with -Update that contact is updated
#   - two or more contacts match (or one email belongs to several): "ambiguous", never written
#   - no email in the row: skipped when a contact with the same name and company exists
# -Update fills only EMPTY fields of the existing contact; add -Overwrite to replace differing
# values too. Notes are always appended under a divider, never replaced.
# Limits: -MaxRows (default 2000) and -MaxFileMB (default 20) refuse larger files.
#
# CSV columns are matched by common header names (Full Name / Name, First Name + Last Name,
# E-mail Address / Email, E-mail 2 Address, Business Phone / Phone, Mobile Phone, Home Phone,
# Company, Job Title / Title, Department, Office, Categories, Notes). -Map overrides or adds:
#   -Map "Work Email=Email; Cell=Mobile; Org=Company"
# Fields: FullName, FirstName, LastName, Email, Email2, Email3, Phone, Mobile, HomePhone,
# Company, JobTitle, Department, Office, Categories, Notes. Files are read as UTF-8.
#
# Usage: .\outlook-contacts-import.ps1 -Path "C:\in\people.csv"
# Usage: .\outlook-contacts-import.ps1 -Path "C:\in\cards.vcf" -Confirm -Json
# Usage: .\outlook-contacts-import.ps1 -Path "C:\in\people.csv" -Map "Cell=Mobile" -Update -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

$headerMap = @{
    "full name" = "FullName"; "name" = "FullName"; "fullname" = "FullName"; "display name" = "FullName"
    "first name" = "FirstName"; "given name" = "FirstName"; "last name" = "LastName"; "surname" = "LastName"; "family name" = "LastName"
    "e-mail address" = "Email"; "email address" = "Email"; "email" = "Email"; "e-mail" = "Email"
    "e-mail 2 address" = "Email2"; "email 2" = "Email2"; "e-mail 3 address" = "Email3"; "email 3" = "Email3"
    "business phone" = "Phone"; "work phone" = "Phone"; "phone" = "Phone"; "mobile phone" = "Mobile"; "mobile" = "Mobile"; "cell" = "Mobile"
    "home phone" = "HomePhone"; "company" = "Company"; "organization" = "Company"; "job title" = "JobTitle"; "title" = "JobTitle"
    "department" = "Department"; "office" = "Office"; "office location" = "Office"; "categories" = "Categories"; "notes" = "Notes"
}
$maxBody = 32KB

# Single-line fields lose control characters; Notes keep line breaks and are capped at 32 KB.
function Clean-Field([string]$Field, [string]$Value) {
    if ($Field -eq "Notes") {
        $v = $Value -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', ''
        if ($v.Length -gt $maxBody) { $v = $v.Substring(0, $maxBody) }
        return $v
    }
    return (($Value -replace '[\x00-\x1F\x7F]', ' ') -replace '\s{2,}', ' ').Trim()
}

function ConvertFrom-QuotedPrintable([string]$Text, [string]$Charset) {
    $bytes = New-Object System.Collections.Generic.List[byte]
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        if ($ch -eq '=' -and $i + 2 -lt $Text.Length -and $Text.Substring($i + 1, 2) -match '^[0-9A-Fa-f]{2}$') {
            $bytes.Add([Convert]::ToByte($Text.Substring($i + 1, 2), 16)); $i += 2
        } else { foreach ($b in [Text.Encoding]::UTF8.GetBytes([string]$ch)) { $bytes.Add($b) } }
    }
    $enc = [Text.Encoding]::UTF8
    if ($Charset) { try { $enc = [Text.Encoding]::GetEncoding($Charset) } catch { } }
    return $enc.GetString($bytes.ToArray())
}

# vCard text value: \, \; \n \\ unescaped.
function Unescape-VValue([string]$v) {
    return ($v -replace '\\[nN]', "`n" -replace '\\,', ',' -replace '\\;', ';' -replace '\\\\', '\')
}

# vCard 2.1/3/4: unfolded lines (and quoted-printable soft breaks), FN, N, EMAIL, TEL (by type,
# "tel:" URIs), ORG (company;department), TITLE, NOTE, CATEGORIES. Structured values are split
# on unescaped ";" before unescaping. A card without END:VCARD, or with nothing usable, is invalid.
function Read-VCards([string]$File) {
    $raw = [IO.File]::ReadAllText($File, [Text.Encoding]::UTF8)
    $raw = $raw -replace "`r`n[ `t]", "" -replace "`n[ `t]", ""          # folded lines
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($l in ($raw -split "`r?`n")) {
        # quoted-printable soft line break: the previous line ends with "="
        if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -match '(?i)QUOTED-PRINTABLE' -and $lines[$lines.Count - 1].EndsWith("=")) {
            $lines[$lines.Count - 1] = $lines[$lines.Count - 1].Substring(0, $lines[$lines.Count - 1].Length - 1) + $l
        } else { $lines.Add($l) }
    }
    $cards = New-Object System.Collections.Generic.List[object]
    $cur = $null; $malformed = 0; $notes = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        if ($line -match '^BEGIN:VCARD') { if ($cur) { $malformed++ }; $cur = [ordered]@{}; continue }
        if ($line -match '^END:VCARD') { if ($cur) { $cards.Add($cur) } else { $malformed++ }; $cur = $null; continue }
        if (-not $cur -or $line -notmatch '^([^:]+):(.*)$') { continue }
        $head = $Matches[1]; $value = $Matches[2]
        $parts = $head.Split(';'); $key = ($parts[0] -replace '^item\d+\.', '').ToUpperInvariant()
        $params = ($parts | Select-Object -Skip 1) -join ';'
        if ($params -match '(?i)ENCODING=QUOTED-PRINTABLE') {
            $cs = if ($params -match '(?i)CHARSET=([^;]+)') { $Matches[1] } else { "utf-8" }
            $value = ConvertFrom-QuotedPrintable $value $cs
        }
        $fields = @($value -split '(?<!\\);' | ForEach-Object { Unescape-VValue $_ })   # structured: split first, then unescape
        $val = Unescape-VValue $value
        switch ($key) {
            "FN" { $cur.FullName = $val }
            "N" { if (-not $cur.Contains("LastName")) { $cur.LastName = $fields[0]; if ($fields.Count -gt 1) { $cur.FirstName = $fields[1] } } }
            "EMAIL" { foreach ($slot in "Email", "Email2", "Email3") { if (-not $cur.Contains($slot)) { $cur[$slot] = $val; break } } }
            "TEL" {
                $num = $val -replace '^(?i)tel:', ''
                $slot = if ($params -match '(?i)cell|mobile') { "Mobile" } elseif ($params -match '(?i)home') { "HomePhone" } else { "Phone" }
                if (-not $cur.Contains($slot)) { $cur[$slot] = $num }
            }
            "ORG" { $cur.Company = $fields[0]; if ($fields.Count -gt 1 -and $fields[1]) { $cur.Department = $fields[1] } }
            "TITLE" { $cur.JobTitle = $val }
            "NOTE" { $cur.Notes = $val }
            "CATEGORIES" { $cur.Categories = (($value -split '(?<!\\),' | ForEach-Object { Unescape-VValue $_ }) -join '; ') }
        }
    }
    if ($cur) { $malformed++ }   # BEGIN without END at the end of the file
    return [ordered]@{ Cards = $cards; Malformed = $malformed }
}

function Read-CsvContacts([string]$File, [hashtable]$Extra) {
    $rows = @(Import-Csv -LiteralPath $File -Encoding UTF8)
    $list = New-Object System.Collections.Generic.List[object]
    if ($rows.Count -eq 0) { return , $list }
    $cols = @{}
    foreach ($h in $rows[0].PSObject.Properties.Name) {
        $k = $h.TrimStart([char]0xFEFF).Trim().ToLowerInvariant()   # a BOM can stick to the first header
        if ($Extra.ContainsKey($k)) { $cols[$h] = $Extra[$k] } elseif ($headerMap.ContainsKey($k)) { $cols[$h] = $headerMap[$k] }
    }
    foreach ($r in $rows) {
        $rec = [ordered]@{}
        foreach ($h in $cols.Keys) { $v = ([string]$r.$h).Trim(); if ($v) { $rec[$cols[$h]] = $v } }
        $list.Add($rec)
    }
    return , $list
}

# Field changes an update would make on an existing contact.
function Get-UpdateChanges($Item, $Fields) {
    $changes = New-Object System.Collections.Generic.List[object]
    foreach ($field in $script:ContactFields.Keys) {
        if (-not $Fields.Contains($field) -or -not $Fields[$field]) { continue }
        $prop = $script:ContactFields[$field]
        $new = [string]$Fields[$field]; $old = [string]$Item.$prop
        if ($field -eq "Notes") {
            if (-not $old.Contains($new)) { $changes.Add([ordered]@{ Field = "Notes"; Old = ""; New = "(appended)"; Mode = "append" }) }
            continue
        }
        if ($field -like "Email*") { continue }   # emails are how the row matched
        if (-not $old.Trim()) { $changes.Add([ordered]@{ Field = $field; Old = ""; New = $new; Mode = "fill" }) }
        elseif ($Overwrite -and $old.Trim() -ne $new) { $changes.Add([ordered]@{ Field = $field; Old = $old; New = $new; Mode = "overwrite" }) }
    }
    return , $changes.ToArray()   # a plain array: @() around a generic List throws "Argument types do not match"
}

try {
    $file = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { Stop-WithError "File not found: $Path" }
    $sizeMB = (Get-Item -LiteralPath $file).Length / 1MB
    if ($sizeMB -gt $MaxFileMB) { Stop-WithError "The file is $([Math]::Round($sizeMB, 2)) MB, over the $MaxFileMB MB limit; nothing was imported." -Hint "Split the file, or raise -MaxFileMB." }
    $ext = [IO.Path]::GetExtension($file).ToLowerInvariant()
    $extra = @{}
    $valid = @($script:ContactFields.Keys) + "FirstName", "LastName"
    foreach ($pair in ($Map -split '[;,]')) {
        if ($pair -notmatch '=') { continue }
        $kv = $pair.Split('=', 2); $field = $kv[1].Trim()
        if ($valid -notcontains $field) { Stop-WithError "Unknown field '$field' in -Map." -Hint "Fields: $($valid -join ', ')" }
        $extra[$kv[0].Trim().ToLowerInvariant()] = $field
    }
    # Assign, never pipe: the readers return their lists comma-wrapped.
    $malformed = 0
    if ($ext -eq ".vcf") { $vc = Read-VCards $file; $incoming = $vc.Cards; $malformed = $vc.Malformed }
    elseif ($ext -eq ".csv") { $incoming = Read-CsvContacts $file $extra }
    else { Stop-WithError "Unsupported file type '$ext'." -Hint "Use a .vcf or .csv file." }
    if ($incoming.Count + $malformed -gt $MaxRows) { Stop-WithError "The file has $($incoming.Count + $malformed) rows, over the -MaxRows $MaxRows limit; nothing was imported." -Hint "Split the file, or raise -MaxRows." }

    $null = Connect-Outlook
    $byEmail = @{}   # lower-case email -> list of contact EntryIDs
    $byNameCo = @{}  # "name|company" -> EntryID
    $all = Get-ContactRecords -Namespace $Namespace
    foreach ($c in $all) {
        foreach ($e in @($c.Email, $c.Email2, $c.Email3)) {
            if (-not $e) { continue }
            $k = $e.Trim().ToLowerInvariant()
            if (-not $byEmail.ContainsKey($k)) { $byEmail[$k] = New-Object System.Collections.Generic.List[string] }
            if (-not $byEmail[$k].Contains($c.EntryID)) { $byEmail[$k].Add($c.EntryID) }
        }
        $byNameCo["$(([string]$c.FullName).Trim().ToLowerInvariant())|$(([string]$c.Company).Trim().ToLowerInvariant())"] = $c.EntryID
    }

    $plan = New-Object System.Collections.Generic.List[object]
    $seenInFile = @{}
    foreach ($rec in $incoming) {
        foreach ($k in @($rec.Keys)) { $rec[$k] = Clean-Field $k ([string]$rec[$k]) }
        if (-not $rec.FullName -and ($rec.FirstName -or $rec.LastName)) { $rec.FullName = (@($rec.FirstName, $rec.LastName) | Where-Object { $_ }) -join " " }
        $emails = @(@($rec.Email, $rec.Email2, $rec.Email3) | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Select-Object -Unique)
        $action = "create"; $target = ""; $why = ""
        if (-not $rec.FullName -and $emails.Count -eq 0) { $action = "invalid" }
        elseif (@($emails | Where-Object { $seenInFile.ContainsKey($_) }).Count -gt 0) { $action = "duplicate in file" }
        elseif ($emails.Count -gt 0) {
            # One id at a time: under 5.1, "$ids += <List>" appends the List as ONE element, and a List
            # passed to GetItemFromID hangs COM marshalling at full CPU.
            $ids = @(); foreach ($e in $emails) { if ($byEmail.ContainsKey($e)) { foreach ($x in $byEmail[$e]) { $ids += [string]$x } } }
            $ids = @($ids | Select-Object -Unique)
            if ($ids.Count -gt 1) { $action = "ambiguous"; $why = "$($ids.Count) existing contacts share these emails" }
            elseif ($ids.Count -eq 1) { $target = [string]$ids[0]; $action = $(if ($Update) { "update" } else { "skip (email exists)" }) }
        } else {
            $key = "$(([string]$rec.FullName).Trim().ToLowerInvariant())|$(([string]$rec.Company).Trim().ToLowerInvariant())"
            if ($byNameCo.ContainsKey($key)) { $action = "skip (name and company exist)" }
        }
        foreach ($e in $emails) { $seenInFile[$e] = 1 }
        $plan.Add([ordered]@{ Action = $action; FullName = [string]$rec.FullName; Email = [string]$rec.Email; Target = $target; Why = $why; Fields = $rec; Changes = @() })
    }
    for ($i = 0; $i -lt $malformed; $i++) { $plan.Add([ordered]@{ Action = "invalid"; FullName = "(malformed card)"; Email = ""; Target = ""; Why = "no END:VCARD"; Fields = [ordered]@{}; Changes = @() }) }

    # Field diff of updates (first -Show rows in the preview; all rows when writing).
    $shown = 0
    foreach ($p in $plan) {
        if ($p.Action -ne "update") { continue }
        $shown++
        if ($shown -gt $Show -and -not $Confirm) { continue }
        # Item in a variable and named parameters: Windows PowerShell 5.1 hangs binding an inline
        # COM call as a positional function argument.
        try { $it = $Namespace.GetItemFromID($p.Target); $p.Changes = Get-UpdateChanges -Item $it -Fields $p.Fields; if ($p.Changes.Count -eq 0) { $p.Action = "unchanged" } }
        catch { $p.Action = "failed: $($_.Exception.Message)" }
    }

    $counts = [ordered]@{ Rows = $plan.Count }
    foreach ($a in "create", "update", "unchanged", "skip (email exists)", "skip (name and company exist)", "ambiguous", "duplicate in file", "invalid") { $counts[$a] = @($plan | Where-Object { $_.Action -eq $a }).Count }
    $todo = $counts["create"] + $counts["update"]
    $rowView = { param($items) @(foreach ($p in ($items | Select-Object -First $Show)) { [ordered]@{ Action = $p.Action; FullName = $p.FullName; Email = $p.Email; Why = $p.Why; Changes = @($p.Changes) } }) }

    if (-not $Json -and -not $Confirm -and (Test-ConfirmRequired -Action "import")) {
        Write-Host "=== IMPORT PREVIEW: $([IO.Path]::GetFileName($file)) ($(if ($Update) { if ($Overwrite) { 'update: overwrite' } else { 'update: fill empty fields' } } else { 'new contacts only' })) ===" -ForegroundColor Yellow
        foreach ($k in $counts.Keys) { if ($counts[$k] -or $k -eq "Rows") { Write-Host ("  {0,-30} {1}" -f $k, $counts[$k]) } }
        foreach ($p in ($plan | Select-Object -First $Show)) {
            Write-Host "  [$($p.Action)] $($p.FullName) $(if ($p.Email) { "<$($p.Email)>" }) $(if ($p.Why) { "($($p.Why))" })" -ForegroundColor Gray
            foreach ($ch in $p.Changes) { Write-Host "      $($ch.Field): $(if ($ch.Mode -eq 'append') { 'notes appended' } else { "'$($ch.Old)' -> '$($ch.New)' ($($ch.Mode))" })" -ForegroundColor DarkGray }
        }
        if ($plan.Count -gt $Show) { Write-Host "  ($($plan.Count - $Show) more rows)" -ForegroundColor DarkGray }
    }
    if ($todo -eq 0) {
        if ($Json) { Write-Json ([ordered]@{ Action = "import"; Result = "nothing to import"; Counts = $counts; Rows = (& $rowView $plan) }) }
        elseif ($Confirm -or -not (Test-ConfirmRequired -Action "import")) { Write-Host "Nothing to import." -ForegroundColor Gray }
        exit 0
    }
    if ((Approve-ContactAction -Action "import" -Confirmed $Confirm -What "the import ($todo contacts)" -Json:$Json) -ne $true) {   # fail closed
        if ($Json) { Write-Json ([ordered]@{ Action = "import"; Result = "preview"; Confirmed = $false; Counts = $counts; Rows = (& $rowView $plan) }) }
        exit 0
    }

    $created = 0; $updated = 0; $failed = 0
    foreach ($p in $plan) {
        if ($p.Action -notin "create", "update") { continue }
        try {
            if ($p.Action -eq "create") {
                $item = $Outlook.CreateItem(2)
                # First/last name before FullName: setting a name part makes Outlook rebuild FullName.
                if ($p.Fields.FirstName) { $item.FirstName = [string]$p.Fields.FirstName }
                if ($p.Fields.LastName) { $item.LastName = [string]$p.Fields.LastName }
                foreach ($field in $script:ContactFields.Keys) {
                    if (-not $p.Fields.Contains($field) -or -not $p.Fields[$field]) { continue }
                    $prop = $script:ContactFields[$field]
                    $item.$prop = [string]$p.Fields[$field]
                }
                $item.Save(); $created++
            } else {
                $item = $Namespace.GetItemFromID($p.Target)
                foreach ($ch in (Get-UpdateChanges -Item $item -Fields $p.Fields)) {
                    $prop = $script:ContactFields[$ch.Field]
                    if ($ch.Mode -eq "append") {
                        $old = ([string]$item.Body).TrimEnd()
                        $item.Body = (($old + "`r`n`r`n----- imported $(Get-Date -Format 'yyyy-MM-dd') -----`r`n" + [string]$p.Fields.Notes).TrimStart())
                    } else { $item.$prop = $ch.New }
                }
                $item.Save(); $updated++
            }
            $p.Target = [string]$item.EntryID
        } catch { $failed++; $p.Action = "failed: $($_.Exception.Message)" }
    }
    if ($Json) {
        Write-Json ([ordered]@{ Action = "import"; Result = $(if ($failed) { "partial" } else { "imported" }); Created = $created; Updated = $updated; Failed = $failed; Counts = $counts
            Rows = @(foreach ($p in ($plan | Select-Object -First $Show)) { [ordered]@{ Action = $p.Action; FullName = $p.FullName; Email = $p.Email; EntryID = $p.Target } }) })
    } else {
        Write-Host "=== IMPORTED: $created created, $updated updated, $failed failed ===" -ForegroundColor $(if ($failed) { "Yellow" } else { "Green" })
        foreach ($p in ($plan | Where-Object { $_.Action -like "failed*" })) { Write-Host "  $($p.FullName): $($p.Action)" -ForegroundColor Red }
    }
    if ($failed -gt 0) { exit 1 }
}
catch {
    Stop-WithError $_
}
