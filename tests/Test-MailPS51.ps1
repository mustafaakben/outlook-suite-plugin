<#
.SYNOPSIS
  Runs the read-only mail scripts under Windows PowerShell 5.1 and PowerShell 7 and compares
  exit codes and output size; find/search must return the same EntryIDs. Prints counts only.
#>
param([string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts"))
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name $detail" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$exe, [string]$script, [string[]]$a) {
    $out = & $exe -NoProfile -File (Join-Path $ScriptsDir "outlook-$script.ps1") @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out; Lines = @($out -split "`n" | Where-Object { $_.Trim() }).Count }
}
function EntryIds([string]$json) { @(($json | ConvertFrom-Json).items | ForEach-Object EntryID) -join "," }

$j = (Run pwsh "find" @("-Days", "14", "-Limit", "1", "-Json")).Out | ConvertFrom-Json
$id = $j.items[0].EntryID

$cases = @(
    @("accounts-list"), @("stores"), @("folders", "-Under", "Inbox"), @("list-categories"),
    @("read", "-Days", "2", "-Limit", "5"), @("search", "-Subject", "the", "-Days", "7"),
    @("search-advanced", "-Folders", "Inbox;Sent", "-Days", "3", "-Limit", "5"),
    @("read-body", "-EntryID", $id), @("read-body", "-EntryID", $id, "-Json"),
    @("conversation", "-EntryID", $id, "-HeadersOnly"), @("links", "-EntryID", $id),
    @("headers", "-EntryID", $id), @("attachment-save", "-EntryID", $id, "-ListOnly"),
    @("reminders"), @("signatures"), @("out-of-office", "-Status"), @("selection")
)
foreach ($c in $cases) {
    $name = $c[0]; $rest = @($c | Select-Object -Skip 1)
    $a = Run powershell $name $rest; $b = Run pwsh $name $rest
    Check "5.1 vs 7: $($c -join ' ' -replace '-EntryID [0-9A-F]+', '-EntryID <id>')" ($a.Exit -eq $b.Exit -and $a.Lines -eq $b.Lines) "(exit $($a.Exit)/$($b.Exit), lines $($a.Lines)/$($b.Lines))"
}
foreach ($f in @(@("find", "-Days", "7", "-Limit", "50", "-Json"), @("find", "-Days", "30", "-UnreadOnly", "-Json"), @("search", "-From", "a", "-Days", "7", "-Json"))) {
    $name = $f[0]; $rest = @($f | Select-Object -Skip 1)
    $a = Run powershell $name $rest; $b = Run pwsh $name $rest
    Check "same EntryIDs: $($f -join ' ')" ($a.Exit -eq 0 -and (EntryIds $a.Out) -eq (EntryIds $b.Out)) "($(@(($b.Out | ConvertFrom-Json).items).Count) items)"
}
$o = & powershell -NoProfile -File (Join-Path $ScriptsDir "outlook-call.ps1") find -Days 2 -Limit 3 2>&1 | Out-String
Check "outlook-call under 5.1 runs directly" ($LASTEXITCODE -eq 0 -and $o -notmatch "host unavailable")
$o = & powershell -NoProfile -File (Join-Path $ScriptsDir "outlook-call.ps1") --ping 2>&1 | Out-String
Check "--ping under 5.1 explains it needs pwsh" ($LASTEXITCODE -eq 1 -and $o -match "PowerShell 7")
$o = & powershell -NoProfile -File (Join-Path $ScriptsDir "outlook-host.ps1") 2>&1 | Out-String
Check "host refuses to start under 5.1" ($LASTEXITCODE -eq 1)

if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
