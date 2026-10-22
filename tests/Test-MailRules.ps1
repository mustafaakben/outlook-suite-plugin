<#
.SYNOPSIS
  Live rules test on one disposable, switched-off rule in a chosen mailbox.

.DESCRIPTION
  Creates "suite-test-rule" (OFF, matching only the subject "[suite-test-never-matches]"),
  lists it, renames it, moves it, runs it on an empty test folder, then deletes it.
  The rule never acts on real mail. Prints PASS/FAIL and timings only.
#>
param(
    [Parameter(Mandatory = $true)][string]$Mailbox,
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts")
)
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
$null = Connect-Outlook

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $out = & pwsh -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out; Ms = $sw.ElapsedMilliseconds }
}
function JsonOf($r) { try { return ($r.Out | ConvertFrom-Json) } catch { return $null } }

$store = Get-RuleStore -Namespace $Namespace -Mailbox $Mailbox
$name = "suite-test-rule"
# Empty test folder in that mailbox's Deleted Items, for -MoveToFolder and -Run.
$del = $null; foreach ($f in $store.GetRootFolder().Folders) { if ($f.Name -in "Deleted Items", "Deleted", "Trash") { $del = $f } }
$tf = $null; foreach ($f in $del.Folders) { if ($f.Name -eq "_suite-test") { $tf = $f } }
if (-not $tf) { $tf = $del.Folders.Add("_suite-test") }
# Note: some account types (for example Outlook.com) refuse move/copy actions in rules made by
# script, so this test uses category + stop actions, which every mailbox accepts.

try {
    $r = Run "outlook-rules-list.ps1" @("-Mailbox", $Mailbox, "-Json")
    $before = (JsonOf $r).total
    Check "rules-list -Json ($($r.Ms) ms)" ($r.Exit -eq 0 -and $null -ne $before)

    $common = @("-Name", $name, "-Mailbox", $Mailbox, "-SubjectContains", "[suite-test-never-matches]", "-ExceptHasAttachment", "-AssignCategory", "suite-test-cat", "-StopProcessing", "-Disabled")
    $r = Run "outlook-rules-create.ps1" $common
    Check "rules-create previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW")
    $r = Run "outlook-rules-create.ps1" ($common + @("-Confirm", "-Json"))
    $j = JsonOf $r
    Check "rules-create -Confirm (save $($j.SaveMs) ms)" ($r.Exit -eq 0 -and $j.Rule.Name -eq $name -and -not $j.Rule.Enabled -and $j.Rule.Exceptions.Count -eq 1 -and $j.Rule.Actions.Count -eq 2)
    $r = Run "outlook-rules-create.ps1" ($common + @("-Confirm"))
    Check "rules-create duplicate name fails without -Force" ($r.Exit -eq 1)

    $r = Run "outlook-rules-list.ps1" @("-Mailbox", $Mailbox, "-Name", $name, "-Detailed")
    Check "rules-list -Detailed shows if/except/then" ($r.Exit -eq 0 -and $r.Out -match "if:" -and $r.Out -match "except:" -and $r.Out -match "then:")

    $r = Run "outlook-rules-manage.ps1" @("-Name", $name, "-Mailbox", $Mailbox, "-NewName", "$name-2")
    Check "rules-manage previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW")
    $r = Run "outlook-rules-manage.ps1" @("-Name", $name, "-Mailbox", $Mailbox, "-NewName", "$name-2", "-Confirm", "-Json")
    Check "rules-manage rename" ($r.Exit -eq 0 -and (JsonOf $r).After.Name -eq "$name-2")
    $r = Run "outlook-rules-manage.ps1" @("-Name", "$name-2", "-Mailbox", $Mailbox, "-MoveTo", "1", "-Confirm", "-Json")
    Check "rules-manage move to position 1" ($r.Exit -eq 0 -and (JsonOf $r).After.Order -eq 1)
    $r = Run "outlook-rules-manage.ps1" @("-Name", "$name-2", "-Mailbox", $Mailbox, "-Run", "-Folder", $tf.FolderPath, "-Confirm")
    Check "rules-manage -Run on an empty test folder" ($r.Exit -eq 0)
    $r = Run "outlook-rules-manage.ps1" @("-Name", "$name-2", "-Mailbox", $Mailbox, "-Delete", "-Confirm")
    $after = (JsonOf (Run "outlook-rules-list.ps1" @("-Mailbox", $Mailbox, "-Json"))).total
    Check "rules-manage -Delete (count back to $before)" ($r.Exit -eq 0 -and $after -eq $before)

    $r = Run "outlook-rules-manage.ps1" @("-Name", "no-such-rule-xyz", "-Mailbox", $Mailbox, "-Disable", "-Confirm")
    Check "rules-manage unknown rule exits 1" ($r.Exit -eq 1)
}
finally {
    try {
        $rules = $store.GetRules(); $changed = $false
        foreach ($n in $name, "$name-2") { for ($i = $rules.Count; $i -ge 1; $i--) { if ($rules.Item($i).Name -eq $n) { $rules.Remove($i); $changed = $true } } }
        if ($changed) { $rules.Save($false) }
    } catch { }
}
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
