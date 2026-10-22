<#
.SYNOPSIS
  Live test for the outlook-notes scripts, on disposable "[suite-test]" notes only.

.DESCRIPTION
  Creates test notes, runs every notes script against them, checks the real Outlook state
  after each step, and prints PASS/FAIL lines and counts only (never note text). Removes the
  test notes at the end, including their copies in Deleted Items.
  Run under both shells:  pwsh -File Test-Notes.ps1   /   powershell -File Test-Notes.ps1 -Shell powershell
#>
param(
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-notes/scripts"),
    [ValidateSet("pwsh", "powershell")]
    [string]$Shell = "pwsh"
)

$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
. (Join-Path $ScriptsDir "_notes.ps1")
$null = Connect-Outlook
# Guard: nothing may be sent. Sent Items and Outbox counts are compared at the end; any increase
# fails the run (a send by someone else during the run also trips it: re-run to confirm).
$sentBefore = $Namespace.GetDefaultFolder(5).Items.Count; $outboxBefore = $Namespace.GetDefaultFolder(4).Items.Count
$savedGate = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"   # force the gate on: a user config that turns it off must never let a test send or delete

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $ErrorActionPreference = "Continue"   # Windows PowerShell 5.1 turns child stderr into terminating errors under Stop
    $out = & $Shell -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}
function RunJson([string]$script, [object[]]$a) {
    $r = Run $script ($a + "-Json")
    try { return [pscustomobject]@{ Exit = $r.Exit; Obj = ($r.Out | ConvertFrom-Json); Out = $r.Out } } catch { return [pscustomobject]@{ Exit = $r.Exit; Obj = $null; Out = $r.Out } }
}
function Item($id) { return $Namespace.GetItemFromID($id) }

$stamp = Get-Date -Format "HHmmss"
$word = "zq$stamp"          # unique search word, placed after the 255-character preview
$created = New-Object System.Collections.Generic.List[string]

try {
    $long = "[suite-test] long note $stamp`r`n" + ("filler text " * 40) + "`r`nhidden $word here. Link: https://example.com/suite-test"
    $r = RunJson "outlook-notes-create.ps1" @("-Body", $long, "-Color", "Pink", "-Categories", "SuiteTest")
    $idLong = $r.Obj.EntryID; if ($idLong) { $created.Add($idLong) }
    Check "create (json, pink, category)" ($r.Exit -eq 0 -and $idLong -and $r.Obj.Color -eq "Pink" -and $r.Obj.Categories -eq "SuiteTest" -and (Item $idLong).Color -eq 2)
    $r = Run "outlook-notes-create.ps1" @("-Body", "[suite-test] short note $stamp", "-Color", "Green")
    $idShort = (Get-NoteRecords -Namespace $Namespace -Search "short note $stamp") | Select-Object -First 1 | ForEach-Object { $_.EntryID }
    if ($idShort) { $created.Add($idShort) }
    Check "create (text output)" ($r.Exit -eq 0 -and $r.Out -match "NOTE CREATED" -and $idShort)

    # --- list / search / filters ---
    $l = RunJson "outlook-notes-list.ps1" @("-Search", $word)
    Check "search finds text beyond the 255-char preview" ($l.Exit -eq 0 -and $l.Obj.total -eq 1 -and $l.Obj.items[0].EntryID -eq $idLong) "total=$($l.Obj.total)"
    $l = RunJson "outlook-notes-list.ps1" @("-Search", $word.ToUpper())
    Check "search ignores case" ($l.Obj.total -eq 1)
    $l = RunJson "outlook-notes-list.ps1" @("-Search", "[suite-test] short note $stamp")
    Check "search with brackets is literal" ($l.Obj.total -eq 1 -and $l.Obj.items[0].EntryID -eq $idShort)
    $l = RunJson "outlook-notes-list.ps1" @("-Search", $stamp, "-Color", "Green")
    Check "color filter" ($l.Obj.total -eq 1 -and $l.Obj.items[0].EntryID -eq $idShort)
    $l = RunJson "outlook-notes-list.ps1" @("-Search", $stamp, "-Category", "SuiteTest")
    Check "category filter" ($l.Obj.total -eq 1 -and $l.Obj.items[0].EntryID -eq $idLong)
    $l = RunJson "outlook-notes-list.ps1" @("-Search", $stamp)
    Check "list json shape and newest first" ($l.Obj.total -eq 2 -and $l.Obj.items[0].Modified -and $l.Obj.items[0].Preview -notmatch "https://")
    $all = RunJson "outlook-notes-list.ps1" @("-Limit", "1000")
    $count = 0; foreach ($i in $Namespace.GetDefaultFolder(12).Items) { if ($i.MessageClass -like "IPM.StickyNote*") { $count++ } }
    Check "note count matches the folder ($count)" ($all.Obj.total -eq $count) "table=$($all.Obj.total)"
    $t = Run "outlook-notes-list.ps1" @("-Search", $stamp)
    Check "list text output" ($t.Exit -eq 0 -and $t.Out -match "OUTLOOK NOTES" -and $t.Out -match "\[Pink\]") "exit=$($t.Exit)"

    # Review fixes: several categories, literal % and _ in searches, JSON delete preview.
    $r = RunJson "outlook-notes-create.ps1" @("-Body", "[suite-test] multi $stamp sale 50%_off$stamp and aXb$stamp", "-Categories", "SuiteTestA, SuiteTestB")
    $idMulti = $r.Obj.EntryID; if ($idMulti) { $created.Add($idMulti) }
    $l = RunJson "outlook-notes-list.ps1" @("-Search", $stamp, "-Category", "SuiteTestB")
    Check "category filter matches one of several" ($l.Obj.total -eq 1 -and $l.Obj.items[0].EntryID -eq $idMulti)
    $l = RunJson "outlook-notes-list.ps1" @("-Search", $stamp, "-Category", "SuiteTes")
    Check "category filter needs the whole name" ($l.Obj.total -eq 0)
    $l = RunJson "outlook-notes-list.ps1" @("-Search", "50%_off$stamp")
    Check "search with % and _ is literal (found)" ($l.Obj.total -eq 1)
    $l = RunJson "outlook-notes-list.ps1" @("-Search", "a_b$stamp")
    Check "search with _ is literal (no wildcard match)" ($l.Obj.total -eq 0)
    $p2 = RunJson "outlook-notes-delete.ps1" @("-EntryID", $idMulti)
    $still = (RunJson "outlook-notes-list.ps1" @("-Search", "multi $stamp")).Obj.total
    Check "delete -Json without -Confirm prints a preview and deletes nothing" ($p2.Exit -eq 0 -and $p2.Obj.Result -eq "preview" -and $p2.Obj.Done -eq $false -and $still -eq 1)

    # --- read ---
    $r = RunJson "outlook-notes-read.ps1" @("-EntryID", $idLong)
    Check "read json: full body, links listed" ($r.Exit -eq 0 -and $r.Obj.Body -match $word -and $r.Obj.Body -match "\[URL-001\]" -and @($r.Obj.Links).Count -eq 1) "exit=$($r.Exit) links=$(@($r.Obj.Links).Count)"
    $r = Run "outlook-notes-read.ps1" @("-Index", "1", "-Search", $word)
    Check "read by -Index with the same filters" ($r.Exit -eq 0 -and $r.Out -match $word)
    $r = Run "outlook-notes-read.ps1" @("-EntryID", $idLong, "-StripLinks", "`$false")
    Check "read -StripLinks `$false keeps the URL" ($r.Exit -eq 0 -and $r.Out -match "https://example.com/suite-test")
    $taskFolderItem = $Namespace.GetDefaultFolder(6)
    $r = Run "outlook-notes-read.ps1" @("-EntryID", "00")
    Check "read with a bad EntryID fails" ($r.Exit -eq 1)

    # --- delete ---
    $r = Run "outlook-notes-delete.ps1" @("-EntryID", $idShort)
    Check "delete without -Confirm is a preview" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY" -and @(Get-NoteRecords -Namespace $Namespace -Search "short note $stamp").Count -eq 1)
    $r = RunJson "outlook-notes-delete.ps1" @("-EntryID", $idShort, "-Confirm")
    # Checked from a separate process: this test session keeps showing an item it has opened.
    $left = (RunJson "outlook-notes-list.ps1" @("-Search", "short note $stamp")).Obj.total
    Check "delete -Confirm" ($r.Exit -eq 0 -and $r.Obj.Result -eq "deleted" -and $left -eq 0) "exit=$($r.Exit) result=$($r.Obj.Result) left=$left"
    if ($r.Exit -eq 0) { [void]$created.Remove($idShort) }
}
finally {
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $savedGate
    # Delete (moves to Deleted Items), then remove those copies by exact subject.
    $del = $Namespace.GetDefaultFolder(3)
    $subjects = New-Object System.Collections.Generic.List[string]
    $subjects.Add("[suite-test] short note $stamp")
    foreach ($id in $created) { try { $x = Item $id; $subjects.Add($x.Subject); $x.Delete() } catch { } }
    $n = 0
    for ($try = 0; $try -lt 10 -and $subjects.Count -gt 0; $try++) {
        foreach ($s in @($subjects)) {
            $hits = @($del.Items.Restrict("[Subject] = '$($s.Replace("'", "''"))'"))
            foreach ($x in $hits) { try { $x.Delete(); $n++ } catch { } }
            if ($hits.Count -gt 0) { [void]$subjects.Remove($s) }
        }
        if ($subjects.Count -gt 0) { Start-Sleep -Seconds 1 }
    }
    Write-Host "Cleanup: $n test item(s) removed$(if ($subjects.Count) { "; not found in Deleted Items: $($subjects.Count)" })." -ForegroundColor Gray
}
$sentAfter = $Namespace.GetDefaultFolder(5).Items.Count; $outboxAfter = $Namespace.GetDefaultFolder(4).Items.Count
Check "nothing sent (Sent Items $sentBefore -> $sentAfter, Outbox $outboxBefore -> $outboxAfter)" ($sentAfter -le $sentBefore -and $outboxAfter -le $outboxBefore)
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASS ($Shell)" -ForegroundColor Green } else { Write-Host "$($script:fails) FAILED ($Shell)" -ForegroundColor Red }
exit $script:fails
