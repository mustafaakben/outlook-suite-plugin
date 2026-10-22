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

# Force the confirm gate on for the whole run, whatever the user's config says, and restore it after.
$prevGate = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"

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
    # Outlook registers a new reminder asynchronously; wait until it shows up (up to 60 s).
    $mine = @(); $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline) {
        $l = JsonOf (Run "outlook-reminders.ps1" @("-Json"))
        $mine = @($l.reminders | Where-Object { $_.Caption -like "*suite-test*" })
        if ($mine.Count) { break }
        Start-Sleep -Seconds 5
    }
    if ($mine.Count -eq 0) { Write-Host "SKIP  reminders List (Outlook had not registered the new test reminder within 60 s; listed $($l.count) others)" -ForegroundColor Yellow }
    else { Check "reminders List shows the test reminder" ($l.count -ge 1) "count=$($l.count)" }
    # Reminders are asynchronous: Outlook pops them up on the desktop and reorders them, and a
    # person can close the pop-up. So the outcome is checked on the task itself, and a reminder
    # that disappears before the script acts is reported as SKIP, not FAIL.
    # (While tests run, leave any "[suite-test] reminder" pop-up alone.)
    $r = Run "outlook-reminders.ps1" @("-Action", "Dismiss", "-Caption", "suite-test")
    if ($r.Out -match "No matching reminder") { Write-Host "SKIP  reminders Dismiss preview (reminder removed by Outlook or the desktop before the script ran)" -ForegroundColor Yellow }
    else { Check "reminders Dismiss previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW") }
    $r = Run "outlook-reminders.ps1" @("-Action", "Dismiss", "-Caption", "suite-test", "-Confirm", "-Json")
    if ($r.Out -match "No matching reminder") { Write-Host "SKIP  reminders Dismiss -Confirm (reminder removed before the script ran)" -ForegroundColor Yellow }
    else {
        $t2 = $Namespace.GetItemFromID($task.EntryID)
        Check "reminders Dismiss -Confirm (task reminder now off)" ($r.Exit -eq 0 -and -not $t2.ReminderSet) "exit=$($r.Exit) reminderSet=$($t2.ReminderSet)"
    }

    # Junk: move a test draft to Junk and back; block/safe sender only as preview
    $r = Run "outlook-junk.ps1" @("-Action", "MarkJunk", "-EntryID", $did, "-Json")
    $jk = JsonOf $r
    Check "junk MarkJunk (test draft into Junk Email)" ($r.Exit -eq 0 -and $Namespace.GetItemFromID($jk.EntryID).Parent.Name -match "Junk")
    $r = Run "outlook-junk.ps1" @("-Action", "NotJunk", "-EntryID", $jk.EntryID, "-Json")
    $nj = JsonOf $r
    Check "junk NotJunk (back to Inbox)" ($r.Exit -eq 0 -and $Namespace.GetItemFromID($nj.EntryID).Parent.Name -eq "Inbox")
    $did = $nj.EntryID
    $r = Run "outlook-junk.ps1" @("-Action", "BlockSender", "-EntryID", $recent[0].EntryID)
    Check "junk BlockSender previews only" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW")
    $r = Run "outlook-junk.ps1" @("-Action", "Lists")
    Check "junk Lists without the Exchange module explains (exit 1)" ($r.Exit -eq 1 -and $r.Out -match "Exchange Online")

    # Signatures: a throwaway signature only
    $sn = "suite-test-sig"
    $r = Run "outlook-signatures.ps1" @("-Action", "Create", "-Name", $sn, "-Text", "Test Signature`nhttps://example.com/sig")
    Check "signatures Create" ($r.Exit -eq 0 -and (Test-Path (Join-Path $env:APPDATA "Microsoft\Signatures\$sn.htm")))
    $r = Run "outlook-signatures.ps1" @("-Action", "List", "-Json")
    Check "signatures List includes it" ($r.Exit -eq 0 -and @((JsonOf $r).Signatures | Where-Object Name -eq $sn).Count -eq 1)
    $r = Run "outlook-signatures.ps1" @("-Action", "Show", "-Name", $sn)
    Check "signatures Show (link as reference)" ($r.Exit -eq 0 -and $r.Out -match "Test Signature" -and $r.Out -match "URL-001")
    $r = Run "outlook-signatures.ps1" @("-Action", "Create", "-Name", $sn, "-Text", "x")
    Check "signatures Create existing without -Force fails" ($r.Exit -eq 1)
    $r = Run "outlook-signatures.ps1" @("-Action", "Insert", "-Name", $sn, "-EntryID", $did)
    Check "signatures Insert into draft" ($r.Exit -eq 0 -and $Namespace.GetItemFromID($did).HTMLBody -match "suite-signature")
    $r = Run "outlook-signatures.ps1" @("-Action", "SetDefault", "-Name", $sn)
    Check "signatures SetDefault previews only" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW")
    $r = Run "outlook-signatures.ps1" @("-Action", "Delete", "-Name", $sn, "-Confirm")
    Check "signatures Delete -Confirm" ($r.Exit -eq 0 -and -not (Test-Path (Join-Path $env:APPDATA "Microsoft\Signatures\$sn.htm")))

    # Stores (read-only)
    $r = Run "outlook-stores.ps1" @("-TopFolders", "5", "-Json")
    $st = JsonOf $r
    Check "stores with sizes and top folders" ($r.Exit -eq 0 -and @($st.stores | Where-Object Default).Count -eq 1 -and @($st.topFolders).Count -eq 5)

    # Saved search folder (created, then removed)
    $r = Run "outlook-search-advanced.ps1" @("-Folders", "Inbox", "-Subject", "zzq-suite-never", "-SaveAs", "suite-test-search", "-Json")
    $sf = $null; foreach ($f in $Namespace.DefaultStore.GetSearchFolders()) { if ($f.Name -eq "suite-test-search") { $sf = $f } }
    Check "search-advanced -SaveAs creates a Search Folder" ($r.Exit -eq 0 -and $sf)
    if ($sf) { $sf.Delete() }
}
finally {
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $prevGate
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    $q = "@SQL=""urn:schemas:httpmail:textdescription"" LIKE '%$marker%' OR ""urn:schemas:httpmail:subject"" LIKE '[suite-test]%'"
    Get-ChildItem (Join-Path $env:APPDATA "Microsoft\Signatures") -Filter "suite-test-sig*" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    foreach ($pass in 1, 2) { foreach ($f in @($Namespace.GetDefaultFolder(16), $Namespace.GetDefaultFolder(13), $Namespace.GetDefaultFolder(6), $Namespace.GetDefaultFolder(23), $Namespace.GetDefaultFolder(3), $sandbox)) { try { foreach ($i in @($f.Items.Restrict($q))) { try { $i.Delete() } catch { } } } catch { } } }
    foreach ($f in $Namespace.DefaultStore.GetSearchFolders()) { if ($f.Name -like "suite-test*") { try { $f.Delete() } catch { } } }
}
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
