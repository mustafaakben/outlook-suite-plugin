<#
.SYNOPSIS
  Checks outlook-smart-reply.ps1 and outlook-reply.ps1 -SmartReply against an independent read of
  Copilot's suggested replies. Creates ONE reply draft (never sent) and deletes it. Prints counts only.
#>
param([string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts"))
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name $detail" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [string[]]$a, [string]$exe = "pwsh") {
    $out = & $exe -NoProfile -File (Join-Path $ScriptsDir "outlook-$script.ps1") @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}

# Independent truth: suggestions read straight from the property, newest 100 Inbox items of 3 days.
$app = New-Object -ComObject Outlook.Application
$tag = "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/EntityExtraction%2FSmartReplyForEmail"
$from = (Get-Date).Date.AddDays(-3)
$items = $app.Session.GetDefaultFolder(6).Items.Restrict("[ReceivedTime] >= '$($from.ToString('MM/dd/yyyy HH:mm'))'")
$items.Sort("[ReceivedTime]", $true)
$with = @(); $without = @(); $k = 0
foreach ($it in $items) {
    if ($k -ge 100) { break }; $k++
    $has = $false; try { $null = $it.PropertyAccessor.GetProperty($tag); $has = $true } catch { }
    if ($has) { $with += $it.EntryID } else { $without += $it.EntryID }
}
Write-Host "Independent: $k items since $($from.ToString('yyyy-MM-dd')), $($with.Count) with suggestions" -ForegroundColor Gray

# List mode
$r = Run "smart-reply" @("-DateFrom", $from.ToString("yyyy-MM-dd"), "-MaxScan", "100", "-Limit", "100", "-Json")
$j = $r.Out | ConvertFrom-Json
Check "list finds every email with suggestions" ($r.Exit -eq 0 -and $j.count -eq $with.Count) "($($j.count)/$($with.Count))"
Check "every listed email has 1 to 3 suggestions" (@($j.items | Where-Object { $_.SmartReplies.Count -lt 1 -or $_.SmartReplies.Count -gt 3 }).Count -eq 0)
$t = Run "smart-reply" @("-DateFrom", $from.ToString("yyyy-MM-dd"), "-MaxScan", "100", "-Limit", "100")
Check "text list exits 0 with a summary line" ($t.Exit -eq 0 -and $t.Out -match "=== \d+ with suggestions")
$p = (Run "smart-reply" @("-DateFrom", $from.ToString("yyyy-MM-dd"), "-MaxScan", "100", "-Limit", "100", "-Json") "powershell").Out | ConvertFrom-Json
Check "same EntryIDs under Windows PowerShell 5.1" ((@($j.items.EntryID) -join ",") -eq (@($p.items.EntryID) -join ",")) "($(@($p.items).Count) items)"

if ($with.Count -eq 0) { Write-Host "No email with suggestions in range; skipping single-email and draft checks." -ForegroundColor Yellow }
else {
    $id = $with[0]
    $one = (Run "smart-reply" @("-EntryID", $id, "-Json")).Out | ConvertFrom-Json
    Check "one email: 1 to 3 suggestions" ($one.SmartReplies.Count -ge 1 -and $one.SmartReplies.Count -le 3) "($($one.SmartReplies.Count))"

    # Draft from suggestion 1: same thread, suggestion text at the top, quoted thread kept. Then delete it.
    $d = Run "smart-reply" @("-EntryID", $id, "-Use", "1", "-Json")
    $draft = $null; try { $draft = $d.Out | ConvertFrom-Json } catch { }
    Check "-Use 1 saves a draft" ($d.Exit -eq 0 -and $draft.EntryID)
    if ($draft.EntryID) {
        $orig = $app.Session.GetItemFromID($id); $dm = $app.Session.GetItemFromID($draft.EntryID)
        Check "draft is unsent and in the same thread" (-not $dm.Sent -and $dm.ConversationID -eq $orig.ConversationID)
        Check "draft body starts with suggestion 1" ($dm.Body.TrimStart().StartsWith($one.SmartReplies[0].Trim()))
        Check "draft keeps the quoted thread" ($dm.Body.Length -gt ($one.SmartReplies[0].Length + 40))
        Check "draft uses Outlook's Aptos style" ($dm.HTMLBody -match "elementToProof")
        $del = Run "delete" @("-EntryID", $draft.EntryID, "-Confirm")
        Check "test draft deleted" ($del.Exit -eq 0)
    }
}

# Errors exit 1
$noSugg = if ($without.Count) { $without[0] } else { "" }
$cases = @(@("smart-reply", @("-Use", "1")), @("reply", @("-EntryID", "0000", "-SmartReply", "1", "-Body", "x")))
if ($noSugg) { $cases += , @("smart-reply", @("-EntryID", $noSugg, "-Use", "1")) }
if ($with.Count) { $cases += , @("reply", @("-EntryID", $with[0], "-SmartReply", "1", "-Body", "x")) }
foreach ($c in $cases) {
    $e = Run $c[0] $c[1]
    Check "error exits 1: $($c[0]) $(($c[1] -join ' ') -replace '[0-9A-F]{40,}', '<id>')" ($e.Exit -eq 1 -and $e.Out -match "Error")
}
$e = Run "smart-reply" @("-EntryID", "0000", "-Use", "4")
Check "-Use 4 is rejected" ($e.Exit -ne 0)

if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
