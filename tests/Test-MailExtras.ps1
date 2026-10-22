<#
.SYNOPSIS
  Live test for selection, headers, conversation actions, reminders, copy, .msg import and
  saved search folders. Real mail is only READ (counts printed); writes happen on test items.
#>
param([string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts"))
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
$null = Connect-Outlook

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $out = & pwsh -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}
function JsonOf($r) { try { return ($r.Out | ConvertFrom-Json) } catch { return $null } }

$marker = "suite-test-marker"
$tmp = Join-Path $env:TEMP "suite-extras-test"; New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$sandbox = $null; foreach ($f in $Namespace.GetDefaultFolder(3).Folders) { if ($f.Name -eq "_suite-test") { $sandbox = $f } }
if (-not $sandbox) { $sandbox = $Namespace.GetDefaultFolder(3).Folders.Add("_suite-test") }

try {
    # --- Read-only on real mail ---
    $r = Run "outlook-selection.ps1" @("-Json")
    Check "selection (read-only)" ($r.Exit -eq 0 -or $r.Out -match "No Outlook window") "exit=$($r.Exit)"

    $recent = (JsonOf (Run "outlook-find.ps1" @("-Days", "14", "-Limit", "30", "-Json"))).items
    $withHeaders = $null
    foreach ($m in $recent) { $h = JsonOf (Run "outlook-headers.ps1" @("-EntryID", $m.EntryID, "-Json")); if ($h -and $h.Hops -gt 0) { $withHeaders = $h; break } }
    Check "headers summary (hops, auth fields parsed)" ($withHeaders -and $withHeaders.Hops -ge 1 -and $withHeaders.MessageID) "found=$([bool]$withHeaders)"
    $r = Run "outlook-headers.ps1" @("-EntryID", $recent[0].EntryID, "-Raw")
    Check "headers -Raw" ($r.Exit -eq 0 -or $r.Out -match "no internet headers")

    $thread = ($recent | Group-Object ConversationID | Where-Object Count -ge 2 | Select-Object -First 1)
    $tid = if ($thread) { $thread.Group[0].EntryID } else { $recent[0].EntryID }
    $r = Run "outlook-conversation-actions.ps1" @("-EntryID", $tid, "-Action", "Status", "-Json")
    $s = JsonOf $r
    Check "conversation Status (read-only)" ($r.Exit -eq 0 -and $s.Messages -ge 1 -and $null -ne $s.AlwaysDelete)
    $r = Run "outlook-conversation-actions.ps1" @("-EntryID", $tid, "-Action", "AlwaysMove", "-Folder", "Deleted Items/_suite-test")
    $s2 = JsonOf (Run "outlook-conversation-actions.ps1" @("-EntryID", $tid, "-Action", "Status", "-Json"))
    Check "conversation AlwaysMove previews only (nothing set)" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and -not $s2.AlwaysMoveTo)

    # --- Writes on test items ---
    $d = $Outlook.CreateItem(0); $d.Subject = "[suite-test] extras"; $d.Body = "$marker extras"; $d.Save(); $did = $d.EntryID

    $r = Run "outlook-copy.ps1" @("-EntryID", $did, "-Folder", "Deleted Items/_suite-test", "-Json")
    $c = JsonOf $r
    Check "copy to folder (new EntryID, original kept)" ($r.Exit -eq 0 -and $c.EntryID -ne $did -and $Namespace.GetItemFromID($c.EntryID).Parent.Name -eq "_suite-test" -and $Namespace.GetItemFromID($did))

    $r = Run "outlook-save-as.ps1" @("-EntryID", $did, "-Path", $tmp, "-Format", "msgunicode", "-Json")
    $msg = (JsonOf $r).File
    $r = Run "outlook-import-msg.ps1" @("-Path", $msg, "-Folder", "Deleted Items/_suite-test", "-Json")
    $im = JsonOf $r
    Check "import .msg (round-trip via save-as)" ($r.Exit -eq 0 -and @($im.Imported).Count -eq 1 -and $Namespace.GetItemFromID($im.Imported[0].EntryID).Subject -eq "[suite-test] extras")

    # Drafts never raise reminders, so use a test task with a reminder that is already due.
    $task = $Outlook.CreateItem(3); $task.Subject = "[suite-test] reminder"; $task.Body = $marker
    $task.ReminderSet = $true; $task.ReminderTime = (Get-Date).AddMinutes(-2); $task.Save()
    Start-Sleep -Seconds 5
    $l = JsonOf (Run "outlook-reminders.ps1" @("-Json"))
    $mine = @($l.reminders | Where-Object { $_.Caption -like "*suite-test*" })
    Check "reminders List shows the test reminder" ($l.count -ge 1 -and $mine.Count -ge 1) "count=$($l.count) mine=$($mine.Count)"
    if ($mine.Count) {
        $r = Run "outlook-reminders.ps1" @("-Action", "Dismiss", "-Caption", "suite-test")
        Check "reminders Dismiss previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW")
        $r = Run "outlook-reminders.ps1" @("-Action", "Dismiss", "-Caption", "suite-test", "-Confirm", "-Json")
        $after = @((JsonOf (Run "outlook-reminders.ps1" @("-Json"))).reminders | Where-Object { $_.Caption -like "*suite-test*" -and $_.Visible })
        Check "reminders Dismiss -Confirm" ($r.Exit -eq 0 -and $after.Count -eq 0) "exit=$($r.Exit)"
    }

    # Saved search folder (created, then removed)
    $r = Run "outlook-search-advanced.ps1" @("-Folders", "Inbox", "-Subject", "zzq-suite-never", "-SaveAs", "suite-test-search", "-Json")
    $sf = $null; foreach ($f in $Namespace.DefaultStore.GetSearchFolders()) { if ($f.Name -eq "suite-test-search") { $sf = $f } }
    Check "search-advanced -SaveAs creates a Search Folder" ($r.Exit -eq 0 -and $sf)
    if ($sf) { $sf.Delete() }
}
finally {
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    $q = "@SQL=""urn:schemas:httpmail:textdescription"" LIKE '%$marker%' OR ""urn:schemas:httpmail:subject"" LIKE '[suite-test]%'"
    foreach ($pass in 1, 2) { foreach ($f in @($Namespace.GetDefaultFolder(16), $Namespace.GetDefaultFolder(13), $Namespace.GetDefaultFolder(3), $sandbox)) { try { foreach ($i in @($f.Items.Restrict($q))) { try { $i.Delete() } catch { } } } catch { } } }
    foreach ($f in $Namespace.DefaultStore.GetSearchFolders()) { if ($f.Name -like "suite-test*") { try { $f.Delete() } catch { } } }
}
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
