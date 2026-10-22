param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Move", "Copy", "MarkRead", "MarkUnread", "Flag", "Unflag", "AddCategory", "RemoveCategory", "ClearCategories", "Importance", "Delete")]
    [string]$Action,
    [string]$Value = "",

    # Which emails: a list of EntryIDs ...
    [string[]]$EntryID = @(),
    [string]$EntryIDFile = "",
    # ... or a search (same filters as outlook-find.ps1)
    [string]$Folder = "",
    [int]$Days = 7,
    [datetime]$DateFrom,
    [datetime]$DateTo,
    [string]$From = "",
    [string]$To = "",
    [string]$Subject = "",
    [string]$BodyContains = "",
    [switch]$UnreadOnly,
    [switch]$ReadOnly,
    [switch]$HasAttachment,
    [switch]$Flagged,
    [string]$Category = "",
    [int]$Limit = 500,

    [switch]$Confirm,
    [switch]$Json
)

# Outlook Bulk: one action on many emails in a single call.
#   Emails: -EntryID id1,id2,... | -EntryIDFile ids.txt (one per line) | a search: -Folder plus
#           any outlook-find filters (-Days, -From, -Subject, -UnreadOnly, -Category, ...).
#   -Action Move|Copy -Value "<folder path>"   MarkRead | MarkUnread | Flag | Unflag
#           AddCategory|RemoveCategory -Value "<name>"   ClearCategories
#           Importance -Value Low|Normal|High   Delete (to Deleted Items)
# Always gated (many emails at once): previews the count and the first emails unless -Confirm.
# Usage: .\outlook-bulk.ps1 -Action MarkRead -Folder Inbox -From "newsletter@" -Days 30 -Confirm
# Usage: .\outlook-bulk.ps1 -Action Move -Value "Archive" -EntryID $ids -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    if ($Action -in "Move", "Copy", "AddCategory", "RemoveCategory", "Importance" -and -not $Value) { Stop-WithError "-Action $Action needs -Value." }
    if ($Action -eq "Importance" -and $Value -notin "Low", "Normal", "High") { Stop-WithError "-Value must be Low, Normal or High." }
    $null = Connect-Outlook

    # Collect targets.
    $ids = @($EntryID | ForEach-Object { $_ -split '[,\s]+' } | Where-Object { $_ })
    if ($EntryIDFile) { $ids += @(Get-Content -LiteralPath $EntryIDFile | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    $targets = New-Object System.Collections.Generic.List[object]
    if ($ids.Count -gt 0) {
        foreach ($id in ($ids | Select-Object -Unique)) { $targets.Add([ordered]@{ EntryID = $id; StoreID = ""; Subject = $null }) }
    } elseif ($Folder) {
        $src = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
        $range = Resolve-DateRange -Days $Days -DateFrom $DateFrom -DateTo $DateTo
        $q = @{ Folder = $src; From = $range[0]; Subject = $Subject; Sender = $From; Recipient = $To; BodyContains = $BodyContains
            UnreadOnly = $UnreadOnly; ReadOnly = $ReadOnly; HasAttachment = $HasAttachment; Flagged = $Flagged; Category = $Category; Limit = $Limit }
        if ($range[1]) { $q.To = $range[1] }
        $found = Find-MailRecords @q
        foreach ($r in $found.Records) { $targets.Add([ordered]@{ EntryID = $r.EntryID; StoreID = $r.StoreID; Subject = $r.Subject }) }
        if ($found.Total -gt $Limit) { Write-Host "Note: $($found.Total) emails match; only the newest $Limit are included (raise -Limit)." -ForegroundColor Yellow }
    } else {
        Stop-WithError "Say which emails: -EntryID, -EntryIDFile, or -Folder with filters."
    }
    if ($targets.Count -eq 0) {
        if ($Json) { Write-Json ([ordered]@{ Action = $Action; Matched = 0; Done = 0 }); exit 0 }
        Write-Host "No emails matched; nothing to do." -ForegroundColor Gray; exit 0
    }

    $dest = $null
    if ($Action -in "Move", "Copy") { $dest = Resolve-OutlookFolder -Namespace $Namespace -Path $Value }

    if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "bulk"))) {
        Write-Host "=== BULK $($Action.ToUpper()) (PREVIEW) ===" -ForegroundColor Yellow
        Write-Host "  Emails: $($targets.Count)$(if ($Value) { "   Value: $Value" })$(if ($dest) { " ($($dest.FolderPath))" })"
        $show = 0
        foreach ($t in $targets) {
            if ($show -ge 10) { Write-Host "  ... and $($targets.Count - 10) more" -ForegroundColor Gray; break }
            $subj = $t.Subject
            if ($null -eq $subj) { try { $subj = (Get-OutlookItem -Namespace $Namespace -EntryID $t.EntryID).Subject } catch { $subj = "(not found)" } }
            Write-Host "  - $subj" -ForegroundColor Gray
            $show++
        }
        Approve-Action -Action "bulk" -Confirmed $false -What "the bulk $Action" | Out-Null
        exit 0
    }

    $ok = 0; $failed = @(); $newIds = @()
    foreach ($t in $targets) {
        try {
            $it = Get-OutlookItem -Namespace $Namespace -EntryID $t.EntryID -StoreID $t.StoreID
            switch ($Action) {
                "Move" { $m = $it.Move($dest); $newIds += [ordered]@{ Old = $t.EntryID; New = $m.EntryID } }
                "Copy" { $c = $it.Copy(); $m = $c.Move($dest); $newIds += [ordered]@{ Source = $t.EntryID; Copy = $m.EntryID } }
                "MarkRead" { if ($it.UnRead) { $it.UnRead = $false; $it.Save() } }
                "MarkUnread" { if (-not $it.UnRead) { $it.UnRead = $true; $it.Save() } }
                "Flag" { try { $it.MarkAsTask(4) } catch { $it.FlagStatus = 2 }; $it.Save() }
                "Unflag" { try { $it.ClearTaskFlag() } catch { }; if ([int]$it.FlagStatus -ne 0) { $it.FlagStatus = 0 }; $it.Save() }
                "AddCategory" {
                    $cur = @(([string]$it.Categories).Split(',', [StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() })
                    if ($cur -notcontains $Value) { $it.Categories = (@($cur) + $Value) -join ", "; $it.Save() }
                }
                "RemoveCategory" {
                    $cur = @(([string]$it.Categories).Split(',', [StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() })
                    if ($cur -contains $Value) { $it.Categories = (@($cur | Where-Object { $_ -ne $Value })) -join ", "; $it.Save() }
                }
                "ClearCategories" { if ($it.Categories) { $it.Categories = ""; $it.Save() } }
                "Importance" { $it.Importance = @{ "Low" = 0; "Normal" = 1; "High" = 2 }[$Value]; $it.Save() }
                "Delete" { $it.Delete() }
            }
            $ok++
        } catch {
            $failed += [ordered]@{ EntryID = $t.EntryID; Error = $_.Exception.Message }
        }
    }

    $res = [ordered]@{ Action = $Action; Value = $Value; Matched = $targets.Count; Done = $ok; Failed = $failed.Count }
    if ($newIds.Count) { $res.NewEntryIDs = $newIds }
    if ($failed.Count) { $res.Failures = $failed }
    if ($Json) { Write-Json $res; if ($failed.Count) { exit 1 } else { exit 0 } }
    Write-Host "=== BULK $($Action.ToUpper()) DONE ===" -ForegroundColor $(if ($failed.Count) { "Yellow" } else { "Green" })
    Write-Host "  Done: $ok of $($targets.Count)$(if ($failed.Count) { "   Failed: $($failed.Count)" })"
    foreach ($f in $failed | Select-Object -First 10) { Write-Host "  Failed: $($f.EntryID.Substring(0,16))... $($f.Error)" -ForegroundColor Red }
    if ($newIds.Count) {
        Write-Host "  New EntryIDs (moving/copying changes them):"
        foreach ($n in $newIds) { Write-Host "    $(@($n.Values)[-1])" }
    }
    if ($failed.Count) { exit 1 }
}
catch {
    Stop-WithError $_
}
