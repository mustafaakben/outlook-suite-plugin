param(
    [string]$Search = "",
    [switch]$AllStores,
    [switch]$Merge,
    [string]$Plan = "",
    [int]$MaxGroups = 10,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Contacts Dedupe: find likely duplicate contacts. Read-only by default.
#
# Same email: contacts sharing an email address (Email, Email2, Email3; Exchange addresses
#   compared as their resolved SMTP). A group is MERGEABLE only when all its contacts are in one
#   mailbox and every duplicate shares at least one name word with the keeper; otherwise it is
#   report-only with the reason (a shared or role address is not one person).
# Same name: same name (case, punctuation and word order ignored) with no shared email. Report only.
#
# Keeper: the most recently modified contact of the group. -Merge fills the keeper's empty fields
# and free email slots from the duplicates, appends their notes under a divider, combines their
# categories, then deletes the duplicates (to Deleted Items). The report lists, per group, every
# contact to be deleted, every conflicting field (kept vs dropped value) and what cannot be carried
# (extra emails, postal addresses, birthday, photo, ...).
#
# Plan lock: the report prints a plan hash (keepers, contacts to delete, filters). -Merge -Confirm
# needs -Plan <hash> and refuses when the contacts changed since the report. At most -MaxGroups
# groups (default 10) are planned per run. Each group is merged on its own; a failed group does not
# stop the others, and the script then exits 1.
#
# Usage: .\outlook-contacts-dedupe.ps1                              # report (prints the plan hash)
# Usage: .\outlook-contacts-dedupe.ps1 -Search "acme" -Json
# Usage: .\outlook-contacts-dedupe.ps1 -Merge -Confirm -Plan 1a2b3c4d5e6f7a8b

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

# "Doe, John" and "john  doe." -> "doe john"
function Get-NameKey([string]$Name) {
    return ((Get-NameWords $Name) -join ' ')
}
function Get-NameWords([string]$Name) {
    return @(($Name.ToLowerInvariant() -replace '[^\p{L}\p{N}\s]', ' ') -split '\s+' | Where-Object { $_.Length -ge 2 } | Sort-Object -Unique)
}
function Get-Emails($Rec) {
    return @(@($Rec.Email, $Rec.Email2, $Rec.Email3) | Where-Object { $_ -and $_ -match '@' } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique)
}
function Get-ShortHash([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))
    return ([BitConverter]::ToString($bytes).Replace("-", "").Substring(0, 16).ToLowerInvariant())
}

# What merging one group does, read from the live items. With -Apply it also writes the keeper
# and deletes the duplicates. $Items[0] is the keeper.
$singleFields = @("FullName", "Phone", "Mobile", "HomePhone", "Company", "JobTitle", "Department", "Office")
function Invoke-GroupMerge {
    param($Items, [switch]$Apply)
    $keep = $Items[0]
    $filled = New-Object System.Collections.Generic.List[string]
    $conflicts = New-Object System.Collections.Generic.List[object]
    $cannot = New-Object System.Collections.Generic.List[string]
    $notes = 0
    $slots = @("Email1Address", "Email2Address", "Email3Address")
    $types = @("Email1AddressType", "Email2AddressType", "Email3AddressType")
    $have = @(for ($s = 0; $s -lt 3; $s++) { $smtp = Resolve-ContactSmtp $Namespace ([string]$keep.($slots[$s])) ([string]$keep.($types[$s])); if ($smtp) { $smtp.ToLowerInvariant() } })
    $freeSlots = @(for ($s = 0; $s -lt 3; $s++) { if (-not [string]$keep.($slots[$s])) { $slots[$s] } })
    $cats = @(([string]$keep.Categories) -split '\s*[;,]\s*' | Where-Object { $_ })
    $body = [string]$keep.Body
    foreach ($dup in ($Items | Select-Object -Skip 1)) {
        $dupName = [string]$dup.FullName
        foreach ($field in $singleFields) {
            $prop = $script:ContactFields[$field]
            $v = ([string]$dup.$prop).Trim(); $k = ([string]$keep.$prop).Trim()
            if (-not $v) { continue }
            if (-not $k) { if ($Apply) { $keep.$prop = $v }; $filled.Add($field) }
            elseif ($k -ne $v) { $conflicts.Add([ordered]@{ Field = $field; Kept = $k; Dropped = $v; From = $dupName }) }
        }
        for ($s = 0; $s -lt 3; $s++) {
            $smtp = Resolve-ContactSmtp $Namespace ([string]$dup.($slots[$s])) ([string]$dup.($types[$s]))   # never a /O= DN
            if (-not $smtp -or $have -contains $smtp.ToLowerInvariant()) { continue }
            if ($freeSlots.Count -gt 0) {
                $slot = $freeSlots[0]; $freeSlots = @($freeSlots | Select-Object -Skip 1)
                if ($Apply) { $keep.$slot = $smtp }
                $have += $smtp.ToLowerInvariant(); $filled.Add("email $smtp")
            } else { $cannot.Add("email $smtp (no free email slot)") }
        }
        foreach ($cat in (([string]$dup.Categories) -split '\s*[;,]\s*' | Where-Object { $_ })) { if ($cats -notcontains $cat) { $cats += $cat; $filled.Add("category $cat") } }
        $dupBody = [string]$dup.Body
        if ($dupBody.Trim()) {
            $notes++
            $body = ($body.TrimEnd() + "`r`n`r`n----- merged from $dupName ($(Get-Date -Format 'yyyy-MM-dd')) -----`r`n" + $dupBody).TrimStart()
        }
        foreach ($p in @(@("BusinessAddress", "business address"), @("HomeAddress", "home address"), @("OtherAddress", "other address"),
                @("WebPage", "web page"), @("IMAddress", "IM address"), @("BusinessFaxNumber", "business fax"), @("PagerNumber", "pager"),
                @("AssistantName", "assistant"), @("Spouse", "spouse"), @("NickName", "nickname"))) {
            try { if (([string]$dup.($p[0])).Trim()) { $cannot.Add("$($p[1]) from $dupName") } } catch { }
        }
        foreach ($p in @(@("Birthday", "birthday"), @("Anniversary", "anniversary"))) {
            try { if (([datetime]$dup.($p[0])).Year -lt 4500) { $cannot.Add("$($p[1]) from $dupName") } } catch { }
        }
        try { if ($dup.HasPicture) { $cannot.Add("photo from $dupName") } } catch { }
    }
    if ($Apply) {
        $keep.Categories = ($cats -join "; ")
        if ($notes -gt 0) { $keep.Body = $body }
        $keep.Save()
        foreach ($dup in ($Items | Select-Object -Skip 1)) { $dup.Delete() }
    }
    # .ToArray() first: piping these generic Lists into @() throws "Argument types do not match".
    return [ordered]@{ Filled = @($filled.ToArray() | Select-Object -Unique); Conflicts = $conflicts.ToArray(); NotesAppended = $notes; CannotCarry = @($cannot.ToArray() | Select-Object -Unique) }
}

function Open-GroupItems($Group) {
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($m in $Group.Contacts) {
        $it = $Namespace.GetItemFromID($m.EntryID, $m.StoreID)
        if (-not $it) { throw "could not open '$($m.FullName)'" }
        $items.Add($it)
    }
    return , $items
}

try {
    $null = Connect-Outlook
    $all = Get-ContactRecords -Namespace $Namespace -AllStores:$AllStores -Search $Search
    $recs = $all.ToArray()   # @($all) throws "Argument types do not match" on this List

    # Email groups: union-find over shared (resolved) addresses.
    $parent = @{}
    for ($i = 0; $i -lt $recs.Count; $i++) { $parent[$i] = $i }
    function Find-Root($i) { while ($parent[$i] -ne $i) { $parent[$i] = $parent[$parent[$i]]; $i = $parent[$i] }; return $i }
    $owner = @{}
    for ($i = 0; $i -lt $recs.Count; $i++) {
        foreach ($e in (Get-Emails $recs[$i])) {
            if ($owner.ContainsKey($e)) { $a = Find-Root $owner[$e]; $b = Find-Root $i; if ($a -ne $b) { $parent[$b] = $a } }
            else { $owner[$e] = $i }
        }
    }
    $byRoot = @{}
    for ($i = 0; $i -lt $recs.Count; $i++) {
        $r = Find-Root $i
        if (-not $byRoot.ContainsKey($r)) { $byRoot[$r] = New-Object System.Collections.Generic.List[int] }
        $byRoot[$r].Add($i)
    }
    $mergeable = New-Object System.Collections.Generic.List[object]
    $reportOnly = New-Object System.Collections.Generic.List[object]
    foreach ($r in $byRoot.Keys) {
        if ($byRoot[$r].Count -lt 2) { continue }
        # Keeper: most recently modified (ties: list order).
        $idx = @($byRoot[$r] | Sort-Object @{ Expression = { $recs[$_].Modified }; Descending = $true }, @{ Expression = { $_ } })
        $members = @(foreach ($k in $idx) { $recs[$k] })
        $g = [ordered]@{ Reason = "same email"; Keep = $members[0].EntryID; Contacts = $members }
        $stores = @($members | ForEach-Object { $_.StoreID } | Select-Object -Unique)
        $keepWords = Get-NameWords $members[0].FullName
        $noShared = @($members | Select-Object -Skip 1 | Where-Object { $w = Get-NameWords $_.FullName; @($w | Where-Object { $keepWords -contains $_ }).Count -eq 0 })
        if ($stores.Count -gt 1) { $g.Reason = "same email, but the contacts are in different mailboxes (merge only within one mailbox)"; $reportOnly.Add($g) }
        elseif ($noShared.Count -gt 0) { $g.Reason = "same email, but no shared name word (a shared or role address?); review by hand"; $reportOnly.Add($g) }
        else { $mergeable.Add($g) }
    }

    # Same name, not linked by email: report only.
    $byName = @{}
    for ($i = 0; $i -lt $recs.Count; $i++) {
        $nk = Get-NameKey ([string]$recs[$i].FullName)
        if (-not $nk) { continue }
        if (-not $byName.ContainsKey($nk)) { $byName[$nk] = New-Object System.Collections.Generic.List[int] }
        $byName[$nk].Add($i)
    }
    $nameGroups = New-Object System.Collections.Generic.List[object]
    foreach ($nk in $byName.Keys) {
        $idx = $byName[$nk]
        if ($idx.Count -lt 2) { continue }
        if (@($idx | ForEach-Object { Find-Root $_ } | Select-Object -Unique).Count -lt 2) { continue }
        $nameGroups.Add([ordered]@{ Reason = "same name only; never merged"; Contacts = @(foreach ($k in $idx) { $recs[$k] }) })
    }

    # The plan: at most -MaxGroups mergeable groups, in keeper-name order, locked by a hash.
    $sorted = @($mergeable | Sort-Object { $_.Contacts[0].FullName }, { $_.Keep })
    $planned = @($sorted | Select-Object -First $MaxGroups)
    $deferred = $sorted.Count - $planned.Count
    $planText = "search=$Search;allstores=$([bool]$AllStores);max=$MaxGroups`n" + ((@($planned | ForEach-Object {
        "K:$($_.Keep)|D:" + ((@($_.Contacts | Select-Object -Skip 1 | ForEach-Object { $_.EntryID }) | Sort-Object) -join ",") }) | Sort-Object) -join "`n")
    $planHash = Get-ShortHash $planText
    $remove = 0; foreach ($g in $planned) { $remove += $g.Contacts.Count - 1 }

    # Details per planned group (conflicts, what cannot be carried), read from the live items.
    foreach ($g in $planned) {
        try { $g.Detail = Invoke-GroupMerge -Items (Open-GroupItems $g) }
        catch { $g.Detail = [ordered]@{ Error = $_.Exception.Message } }
    }

    $applying = $Merge -and ($Confirm -or -not (Test-ConfirmRequired -Action "delete"))
    if (-not $applying) {
        if ($Json) {
            $o = [ordered]@{}
            if ($Merge) { $o.Action = "merge"; $o.Result = "preview"; $o.Confirmed = $false }
            $o.Contacts = $recs.Count; $o.Plan = $planHash; $o.Groups = $planned.Count; $o.Duplicates = $remove; $o.DuplicateGroups = $planned
            $o.DeferredGroups = $deferred; $o.ReportOnlyGroups = $reportOnly; $o.NameOnlyGroups = $nameGroups.Count; $o.SameNameGroups = $nameGroups
        }
        if (-not $Json) {
            $scope = if ($Search) { " matching '$Search'" } else { "" }
            $title = if ($Merge) { "MERGE DUPLICATES (PREVIEW)" } else { "DUPLICATE CONTACTS" }
            Write-Host "=== ${title}: $($recs.Count) contacts$scope ===" -ForegroundColor $(if ($Merge) { "Yellow" } else { "Cyan" })
            Write-Host "Mergeable (same email, same mailbox, shared name word): $($sorted.Count) groups, planned $($planned.Count), $remove contacts to delete. Report only: $($reportOnly.Count) same-email, $($nameGroups.Count) same-name."
            Write-Host "Keeper rule: the most recently modified contact of each group." -ForegroundColor Gray
            $n = 0
            foreach ($g in $planned) {
                $n++
                Write-Host ""
                Write-Host "Group $n (merge):" -ForegroundColor Yellow
                Write-Host "  keep   $((@($g.Contacts[0].FullName, $g.Contacts[0].Email) | Where-Object { $_ }) -join ' <')$(if ($g.Contacts[0].Email) { '>' }) | modified $($g.Contacts[0].Modified)"
                foreach ($m in ($g.Contacts | Select-Object -Skip 1)) { Write-Host "  delete $($m.FullName)$(if ($m.Email) { " <$($m.Email)>" }) | modified $($m.Modified)" }
                if ($g.Detail.Error) { Write-Host "  cannot read this group: $($g.Detail.Error)" -ForegroundColor Red; continue }
                if ($g.Detail.Filled.Count) { Write-Host "  fills: $($g.Detail.Filled -join ', ')" -ForegroundColor Gray }
                foreach ($x in $g.Detail.Conflicts) { Write-Host "  conflict $($x.Field): keeps '$($x.Kept)', drops '$($x.Dropped)' (from $($x.From))" -ForegroundColor Magenta }
                if ($g.Detail.NotesAppended) { Write-Host "  notes: $($g.Detail.NotesAppended) appended under a divider" -ForegroundColor Gray }
                if ($g.Detail.CannotCarry.Count) { Write-Host "  not carried over: $($g.Detail.CannotCarry -join '; ')" -ForegroundColor DarkYellow }
            }
            foreach ($g in @($reportOnly) + @($nameGroups)) {
                $n++
                Write-Host ""
                Write-Host "Group $n (report only: $($g.Reason)):" -ForegroundColor DarkGray
                foreach ($m in $g.Contacts) { Write-Host "         $($m.FullName)$(if ($m.Email) { " <$($m.Email)>" })" -ForegroundColor DarkGray }
            }
            if ($deferred -gt 0) { Write-Host ""; Write-Host "$deferred more mergeable groups are not in this plan (-MaxGroups $MaxGroups); run again after merging." -ForegroundColor Gray }
            if ($planned.Count -gt 0) {
                $cmd = "-Merge -Confirm -Plan $planHash"
                if ($Search) { $cmd += ' -Search "' + $Search + '"' }
                if ($AllStores) { $cmd += " -AllStores" }
                if ($MaxGroups -ne 10) { $cmd += " -MaxGroups $MaxGroups" }
                Write-Host ""
                Write-Host "Plan: $planHash   To merge exactly this: $cmd" -ForegroundColor Cyan
            }
        }
    }

    if (-not $Merge) {
        if ($Json) { Write-Json $o }
        exit 0
    }
    if ($planned.Count -eq 0) {
        if ($Json) { Write-Json ([ordered]@{ Action = "merge"; Result = "nothing to merge"; Groups = 0; ReportOnlyGroups = $reportOnly.Count; NameOnlyGroups = $nameGroups.Count }) }
        else { Write-Host "Nothing to merge." -ForegroundColor Gray }
        exit 0
    }
    if ((Approve-ContactAction -Action "delete" -Confirmed $Confirm -What "the merge ($remove contacts deleted)" -Json:$Json) -ne $true) {   # fail closed
        if ($Json) { Write-Json $o }
        exit 0
    }
    if (-not $Plan) { Stop-WithError "-Merge needs -Plan <hash> from the report, so it merges exactly what you reviewed." -Hint "Run outlook-contacts-dedupe.ps1 (same -Search/-AllStores/-MaxGroups) and pass the Plan it prints." }
    if ($Plan.Trim().ToLowerInvariant() -ne $planHash) { Stop-WithError "The plan changed since the report (plan $Plan, now $planHash); nothing was merged." -Hint "Run the report again, review it, and pass the new Plan." }

    $results = New-Object System.Collections.Generic.List[object]
    $failed = 0
    foreach ($g in $planned) {
        try {
            $items = Open-GroupItems $g
            $d = Invoke-GroupMerge -Items $items -Apply
            $results.Add([ordered]@{ Status = "merged"; Kept = [string]$items[0].FullName; EntryID = [string]$items[0].EntryID; Deleted = $items.Count - 1
                Filled = $d.Filled; Conflicts = $d.Conflicts; NotesAppended = $d.NotesAppended; CannotCarry = $d.CannotCarry })
        } catch {
            $failed++
            $results.Add([ordered]@{ Status = "failed"; Kept = [string]$g.Contacts[0].FullName; EntryID = [string]$g.Keep; Error = $_.Exception.Message })
        }
    }
    $merged = @($results | Where-Object { $_.Status -eq "merged" })
    $deleted = 0; foreach ($x in $merged) { $deleted += $x.Deleted }
    $untouched = $deferred + $reportOnly.Count + $nameGroups.Count
    if ($Json) {
        Write-Json ([ordered]@{ Action = "merge"; Result = $(if ($failed) { "partial" } else { "merged" }); Plan = $planHash; Groups = $merged.Count; Failed = $failed
            Untouched = $untouched; Deleted = $deleted; DeletedTo = "Deleted Items"; Merged = $results })
    } else {
        Write-Host "=== MERGED: $($merged.Count) groups, $deleted contacts deleted (they are in Deleted Items); failed $failed; untouched $untouched ===" -ForegroundColor $(if ($failed) { "Yellow" } else { "Green" })
        foreach ($x in $results) {
            if ($x.Status -eq "failed") { Write-Host "  FAILED $($x.Kept): $($x.Error)" -ForegroundColor Red; continue }
            Write-Host "  kept $($x.Kept); deleted $($x.Deleted); filled: $(if ($x.Filled.Count) { $x.Filled -join ', ' } else { 'nothing' })"
            foreach ($c in $x.Conflicts) { Write-Host "    kept $($c.Field) '$($c.Kept)', dropped '$($c.Dropped)'" -ForegroundColor Magenta }
            if ($x.CannotCarry.Count) { Write-Host "    not carried over: $($x.CannotCarry -join '; ')" -ForegroundColor DarkYellow }
            Write-Host "    EntryID: $($x.EntryID)" -ForegroundColor DarkGray
        }
    }
    if ($failed -gt 0) { exit 1 }
}
catch {
    Stop-WithError $_
}
