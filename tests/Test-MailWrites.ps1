<#
.SYNOPSIS
  Live write test for the mail action scripts, on a disposable test draft only.

.DESCRIPTION
  Uses the folder "Deleted Items/_suite-test" (created once, kept; out of sight) and a test draft (subject "[suite-test] ..."),
  runs each action script against that draft, checks the real Outlook state after each
  step, and prints PASS/FAIL lines only. Nothing is ever sent, printed or recalled:
  outward-facing scripts are run in preview mode. Cleans up at the end.
#>
param(
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts")
)

$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
$null = Connect-Outlook

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $out = & pwsh -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}
function Item($id) { return $Namespace.GetItemFromID($id) }

# Sandbox
# Sandbox lives in Deleted Items and is kept: Exchange does not let scripts delete folders there,
# so creating and deleting one per run would pile them up.
$deletedRoot = $Namespace.GetDefaultFolder(3)
$sandbox = $null
foreach ($f in $deletedRoot.Folders) { if ($f.Name -eq "_suite-test") { $sandbox = $f } }
if (-not $sandbox) { $sandbox = $deletedRoot.Folders.Add("_suite-test") }
$draft = $Outlook.CreateItem(0)
$draft.Subject = "[suite-test] write test $(Get-Date -Format s)"
$draft.Body = "Test draft created by Test-MailWrites.ps1. Safe to delete. Link: https://example.com/test-link"
$draft.Save()
$id = $draft.EntryID
Write-Host "Sandbox ready; test draft created." -ForegroundColor Gray

try {
    # Flag / unflag
    $r = Run "outlook-flag.ps1" @("-EntryID", $id, "-When", "Tomorrow", "-ReminderDate", (Get-Date).AddDays(2).ToString("yyyy-MM-dd 09:00"))
    $it = Item $id; Check "flag -When Tomorrow" ($r.Exit -eq 0 -and $it.FlagStatus -eq 2 -and $it.ReminderSet) "exit=$($r.Exit) status=$($it.FlagStatus)"
    $r = Run "outlook-unflag.ps1" @("-EntryID", $id, "-Complete")
    $it = Item $id; Check "unflag -Complete on a draft fails clearly (Outlook rule)" ($r.Exit -eq 1 -and $r.Out -match "draft" -and $it.FlagStatus -eq 2)
    $r = Run "outlook-unflag.ps1" @("-EntryID", $id)
    $it = Item $id; Check "unflag (clear)" ($r.Exit -eq 0 -and $it.FlagStatus -eq 0)

    # Read state
    $r = Run "outlook-mark-read.ps1" @("-EntryID", $id, "-Unread")
    $it = Item $id; Check "mark-read -Unread" ($r.Exit -eq 0 -and $it.UnRead)
    $r = Run "outlook-mark-read.ps1" @("-EntryID", $id)
    $it = Item $id; Check "mark-read" ($r.Exit -eq 0 -and -not $it.UnRead)

    # Importance / sensitivity
    $r = Run "outlook-set-importance.ps1" @("-EntryID", $id, "-Importance", "High")
    $it = Item $id; Check "set-importance High" ($r.Exit -eq 0 -and $it.Importance -eq 2)
    $r = Run "outlook-set-sensitivity.ps1" @("-EntryID", $id, "-Sensitivity", "Confidential")
    $it = Item $id; Check "set-sensitivity Confidential" ($r.Exit -eq 0 -and $it.Sensitivity -eq 3)

    # Categories
    $r = Run "outlook-assign-category.ps1" @("-EntryID", $id, "-Category", "suite-test-A,suite-test-B")
    $it = Item $id; Check "assign-category two" ($r.Exit -eq 0 -and $it.Categories -match "suite-test-A" -and $it.Categories -match "suite-test-B")
    $r = Run "outlook-assign-category.ps1" @("-EntryID", $id, "-Category", "suite-test-A")
    $it = Item $id; Check "assign-category idempotent" ($r.Exit -eq 0 -and ([regex]::Matches($it.Categories, "suite-test-A")).Count -eq 1)
    $r = Run "outlook-remove-category.ps1" @("-EntryID", $id, "-Category", "suite-test-A")
    $it = Item $id; Check "remove-category one" ($r.Exit -eq 0 -and $it.Categories -notmatch "suite-test-A" -and $it.Categories -match "suite-test-B")
    $r = Run "outlook-remove-category.ps1" @("-EntryID", $id, "-All")
    $it = Item $id; Check "remove-category -All" ($r.Exit -eq 0 -and -not $it.Categories)

    # Move to sandbox (EntryID may change) and back
    $r = Run "outlook-move.ps1" @("-EntryID", $id, "-Folder", "Deleted Items/_suite-test", "-Json")
    $j = $r.Out | ConvertFrom-Json
    $it = Item $j.EntryID; Check "move to Deleted Items/_suite-test" ($r.Exit -eq 0 -and $it.Parent.Name -eq "_suite-test")
    $id = $j.EntryID
    $r = Run "outlook-move.ps1" @("-EntryID", $id, "-Folder", "No/Such/Folder")
    Check "move to missing folder fails (exit 1)" ($r.Exit -eq 1 -and $r.Out -match "Error:")

    # Save as
    $tmp = Join-Path $env:TEMP "suite-test-save"
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    $r = Run "outlook-save-as.ps1" @("-EntryID", $id, "-Format", "txt", "-Path", $tmp, "-Json")
    $file = ($r.Out | ConvertFrom-Json).File
    Check "save-as txt" ($r.Exit -eq 0 -and (Test-Path -LiteralPath $file))
    $r2 = Run "outlook-save-as.ps1" @("-EntryID", $id, "-Format", "txt", "-Path", $file, "-Json")
    $file2 = ($r2.Out | ConvertFrom-Json).File
    Check "save-as does not overwrite" ($r2.Exit -eq 0 -and $file2 -ne $file -and (Test-Path -LiteralPath $file2))
    Remove-Item -Recurse -Force $tmp

    # Links on our own draft
    $r = Run "outlook-links.ps1" @("-EntryID", $id, "-Ref", "URL-001")
    Check "links -Ref URL-001" ($r.Exit -eq 0 -and $r.Out -match "https://example.com/test-link")

    # Outward-facing: preview only, nothing happens
    $r = Run "outlook-print.ps1" @("-EntryID", $id)
    Check "print previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY")
    $r = Run "outlook-recall.ps1" @("-EntryID", $id, "-DeleteUnread")
    $it = Item $id; Check "recall -DeleteUnread previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY" -and $it)

    # Delete: preview, then confirm, then the global switch
    $r = Run "outlook-delete.ps1" @("-EntryID", $id)
    $still = $true; try { $null = (Item $id).Subject } catch { $still = $false }
    Check "delete previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY" -and $still)

    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "0"
    $r = Run "outlook-delete.ps1" @("-EntryID", $id)
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $null
    $deleted = $Namespace.GetDefaultFolder(3)
    $inDeleted = $null; foreach ($d in $deleted.Items.Restrict("[Subject] = '$($draft.Subject.Replace("'", "''"))'")) { $inDeleted = $d }
    Check "confirmation switched off: delete acts without -Confirm" ($r.Exit -eq 0 -and $inDeleted)

    $r = Run "outlook-delete.ps1" @("-EntryID", $inDeleted.EntryID, "-Permanent", "-Confirm")
    $gone = $true; foreach ($d in $deleted.Items.Restrict("[Subject] = '$($draft.Subject.Replace("'", "''"))'")) { $gone = $false }
    Check "delete -Permanent -Confirm purges" ($r.Exit -eq 0 -and $gone)

    # Master categories (test names only)
    $cat = "suite-test-cat"
    $r = Run "outlook-create-category.ps1" @("-Name", $cat, "-Color", "Green")
    $c = Get-OutlookCategory $Namespace $cat; Check "create-category by color name" ($r.Exit -eq 0 -and $c -and $c.Color -eq 5)
    $r = Run "outlook-create-category.ps1" @("-Name", $cat)
    Check "create-category duplicate fails (exit 1)" ($r.Exit -eq 1)
    $r = Run "outlook-update-category.ps1" @("-Name", $cat, "-NewName", "$cat-2", "-Color", "DarkBlue")
    $c = Get-OutlookCategory $Namespace "$cat-2"; Check "update-category rename + recolor" ($r.Exit -eq 0 -and $c -and $c.Color -eq 23)
    $r = Run "outlook-delete-category.ps1" @("-Name", "$cat-2")
    Check "delete-category previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and (Get-OutlookCategory $Namespace "$cat-2"))
    $r = Run "outlook-delete-category.ps1" @("-Name", "$cat-2", "-Confirm")
    Check "delete-category -Confirm" ($r.Exit -eq 0 -and -not (Get-OutlookCategory $Namespace "$cat-2"))

    # Read-only info scripts
    $r = Run "outlook-accounts-list.ps1" @("-Json"); $j = $r.Out | ConvertFrom-Json
    Check "accounts-list -Json (one default)" ($r.Exit -eq 0 -and $j.count -ge 1 -and @($j.accounts | Where-Object Default).Count -eq 1)
    $r = Run "outlook-list-categories.ps1" @("-Json")
    Check "list-categories -Json" ($r.Exit -eq 0 -and ($r.Out | ConvertFrom-Json).count -ge 0)
    $r = Run "outlook-out-of-office.ps1" @("-Status", "-Json")
    Check "out-of-office -Status (read-only)" ($r.Exit -eq 0 -and ($r.Out | ConvertFrom-Json).AutomaticReplies -in @("On", "Off"))
    $r = Run "outlook-out-of-office.ps1" @("-Enable")
    Check "out-of-office -Enable previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW")

    # Errors
    $r = Run "outlook-flag.ps1" @("-EntryID", "00AB")
    Check "bad EntryID exits 1" ($r.Exit -eq 1 -and $r.Out -match "Error:")
}
finally {
    foreach ($n in "suite-test-cat", "suite-test-cat-2") {
        $i = 0; foreach ($c in $Namespace.Categories) { $i++; if ($c.Name -eq $n) { $Namespace.Categories.Remove($i); break } }
    }
    # Remove every test item this or an earlier interrupted run left behind (subject "[suite-test] ...").
    $mine = "@SQL=""urn:schemas:httpmail:subject"" LIKE '[suite-test]%'"
    foreach ($pass in 1, 2) {   # pass 2 purges what pass 1 moved into Deleted Items
        foreach ($f in @($sandbox, $Namespace.GetDefaultFolder(16), $Namespace.GetDefaultFolder(6), $Namespace.GetDefaultFolder(3))) {
            try { foreach ($i in @($f.Items.Restrict($mine))) { try { $i.Delete() } catch { } } } catch { }
        }
    }
}

Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
