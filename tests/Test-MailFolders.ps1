<#
.SYNOPSIS
  Live test for outlook-folder-manage and outlook-bulk on test folders and test drafts only.

.DESCRIPTION
  Reuses one folder "_suite-fold": brought back from Deleted Items (or created under Drafts),
  renamed, filled with test drafts via bulk move, bulk-marked, emptied and deleted again, so
  at most one empty test folder ever remains in Deleted Items (Exchange does not let scripts
  purge folders there). Prints PASS/FAIL lines only.
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
function Child($parent, $name) { foreach ($c in $parent.Folders) { if ($c.Name -eq $name) { return $c } }; return $null }

$drafts = $Namespace.GetDefaultFolder(16); $deleted = $Namespace.GetDefaultFolder(3)
$marker = "suite-test-marker"
$ids = @()
try {
    # Bring the test folder back, or create it.
    $parked = Child $deleted "_suite-fold"
    if ($parked) {
        $r = Run "outlook-folder-manage.ps1" @("-Action", "Move", "-Folder", "Deleted Items/_suite-fold", "-To", "Drafts", "-Confirm")
        Check "folder Move (back from Deleted Items)" ($r.Exit -eq 0 -and (Child $drafts "_suite-fold"))
    } else {
        $r = Run "outlook-folder-manage.ps1" @("-Action", "Create", "-Folder", "Drafts", "-Name", "_suite-fold")
        Check "folder Create" ($r.Exit -eq 0 -and (Child $drafts "_suite-fold"))
    }
    $r = Run "outlook-folder-manage.ps1" @("-Action", "Create", "-Folder", "Drafts", "-Name", "_suite-fold")
    Check "folder Create duplicate fails" ($r.Exit -eq 1)

    $r = Run "outlook-folder-manage.ps1" @("-Action", "Rename", "-Folder", "Drafts/_suite-fold", "-Name", "_suite-fold2")
    Check "folder Rename previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and (Child $drafts "_suite-fold"))
    $r = Run "outlook-folder-manage.ps1" @("-Action", "Rename", "-Folder", "Drafts/_suite-fold", "-Name", "_suite-fold2", "-Confirm")
    Check "folder Rename -Confirm" ($r.Exit -eq 0 -and (Child $drafts "_suite-fold2") -and -not (Child $drafts "_suite-fold"))
    $null = Run "outlook-folder-manage.ps1" @("-Action", "Rename", "-Folder", "Drafts/_suite-fold2", "-Name", "_suite-fold", "-Confirm")
    $tf = Child $drafts "_suite-fold"

    # Three test drafts (unread) for bulk.
    foreach ($n in 1..3) { $d = $Outlook.CreateItem(0); $d.Subject = "[suite-test] bulk $n"; $d.Body = $marker; $d.Save(); $d.UnRead = $true; $d.Save(); $ids += $d.EntryID }

    $r = Run "outlook-bulk.ps1" @("-Action", "Move", "-Value", "Drafts/_suite-fold", "-EntryID", ($ids -join ","))
    Check "bulk Move previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and $tf.Items.Count -eq 0)
    $r = Run "outlook-bulk.ps1" @("-Action", "Move", "-Value", "Drafts/_suite-fold", "-EntryID", ($ids -join ","), "-Confirm", "-Json")
    $j = JsonOf $r
    Check "bulk Move -Confirm (3 moved, new EntryIDs)" ($r.Exit -eq 0 -and $j.Done -eq 3 -and $tf.Items.Count -eq 3 -and @($j.NewEntryIDs).Count -eq 3)
    $ids = @($j.NewEntryIDs | ForEach-Object { $_.New })

    $r = Run "outlook-bulk.ps1" @("-Action", "AddCategory", "-Value", "suite-test-cat", "-Folder", "Drafts/_suite-fold", "-Days", "1", "-Confirm", "-Json")
    $cats = @(foreach ($i in $tf.Items) { $i.Categories })
    Check "bulk AddCategory by search (-Folder)" ($r.Exit -eq 0 -and (JsonOf $r).Done -eq 3 -and @($cats | Where-Object { $_ -match "suite-test-cat" }).Count -eq 3)
    $r = Run "outlook-bulk.ps1" @("-Action", "Importance", "-Value", "High", "-EntryID", ($ids -join ","), "-Confirm")
    Check "bulk Importance High" ($r.Exit -eq 0 -and @(foreach ($i in $tf.Items) { if ($i.Importance -eq 2) { 1 } }).Count -eq 3)

    $r = Run "outlook-folder-manage.ps1" @("-Action", "MarkAllRead", "-Folder", "Drafts/_suite-fold", "-Confirm", "-Json")
    Check "folder MarkAllRead" ($r.Exit -eq 0 -and $tf.UnReadItemCount -eq 0)
    $r = Run "outlook-bulk.ps1" @("-Action", "MarkUnread", "-EntryID", ($ids -join ","), "-Confirm")
    $unreadNow = @(foreach ($id in $ids) { if ($Namespace.GetItemFromID($id).UnRead) { 1 } }).Count
    Check "bulk MarkUnread (per-item state)" ($r.Exit -eq 0 -and $unreadNow -eq 3) "exit=$($r.Exit) unread=$unreadNow"

    $r = Run "outlook-bulk.ps1" @("-Action", "Move", "-Value", "Drafts/_suite-fold", "-EntryID", "00AB,00CD", "-Confirm", "-Json")
    Check "bulk with bad EntryIDs reports failures (exit 1)" ($r.Exit -eq 1 -and (JsonOf $r).Failed -eq 2)

    $r = Run "outlook-folder-manage.ps1" @("-Action", "Empty", "-Folder", "Drafts/_suite-fold")
    Check "folder Empty previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and $tf.Items.Count -eq 3)
    $r = Run "outlook-folder-manage.ps1" @("-Action", "Empty", "-Folder", "Drafts/_suite-fold", "-Confirm", "-Json")
    Check "folder Empty -Confirm" ($r.Exit -eq 0 -and $tf.Items.Count -eq 0)

    $r = Run "outlook-folder-manage.ps1" @("-Action", "Delete", "-Folder", "Drafts/_suite-fold", "-Confirm")
    Check "folder Delete -Confirm (to Deleted Items)" ($r.Exit -eq 0 -and -not (Child $drafts "_suite-fold") -and (Child $deleted "_suite-fold"))
    $r = Run "outlook-folder-manage.ps1" @("-Action", "Rename", "-Folder", "Drafts/_suite-fold", "-Name", "x", "-Confirm")
    Check "folder action on a missing folder exits 1" ($r.Exit -eq 1)
}
finally {
    $q = "@SQL=""urn:schemas:httpmail:textdescription"" LIKE '%$marker%' OR ""urn:schemas:httpmail:subject"" LIKE '[suite-test]%'"
    foreach ($pass in 1, 2) {
        $folders = @($drafts, $deleted)
        foreach ($n in "_suite-fold", "_suite-fold2") { foreach ($p in $drafts, $deleted) { $c = Child $p $n; if ($c) { $folders += $c } } }
        foreach ($f in $folders) { try { foreach ($i in @($f.Items.Restrict($q))) { try { $i.Delete() } catch { } } } catch { } }
    }
    foreach ($n in "_suite-fold", "_suite-fold2") { $c = Child $drafts $n; if ($c) { try { $c.Delete() } catch { } } }
    $i = 0; foreach ($c in $Namespace.Categories) { $i++; if ($c.Name -eq "suite-test-cat") { $Namespace.Categories.Remove($i); break } }
}
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
