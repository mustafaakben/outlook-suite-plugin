<#
.SYNOPSIS
  Live test of the warm host (outlook-host.ps1 + outlook-call.ps1) with the current scripts,
  then a direct-vs-host benchmark. Prints PASS/FAIL lines and timings only.
#>
param([string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts"), [int]$Runs = 3)
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
$call = Join-Path $ScriptsDir "outlook-call.ps1"

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
# Each client call runs in its own pwsh, the way an agent's shell would.
function ViaHost([object[]]$a) {
    $out = & pwsh -NoProfile -File $call @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}

# Fresh host so it runs the current code.
& pwsh -NoProfile -File $call --stop | Out-Null
Start-Sleep -Seconds 1

$r = ViaHost @("accounts-list")
Check "first call starts the host" ($r.Exit -eq 0 -and $r.Out -match "OUTLOOK ACCOUNTS")
$r = ViaHost @("--ping")
Check "host answers ping" ($r.Exit -eq 0 -and $r.Out -match "pong")

$r = ViaHost @("read-body", "-EntryID", "00AB")
Check "error exits 1 through the host" ($r.Exit -eq 1 -and $r.Out -match "Error:")
Check "host survives a failing script" ((ViaHost @("--ping")).Exit -eq 0)

$r = ViaHost @("find", "-Days", "3", "-Limit", "5", "-Json")
$j = $null; try { $j = $r.Out | ConvertFrom-Json } catch { }
Check "JSON comes back intact" ($r.Exit -eq 0 -and $j -and $j.count -ge 0)

$r = ViaHost @("delete", "-EntryID", $j.items[0].EntryID)
Check "gated preview exits 0 and acts on nothing" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW ONLY")

$r = ViaHost @("find", "-Folder", "No/Such/Folder")
Check "unknown folder exits 1 (no Inbox fallback)" ($r.Exit -eq 1)

# Five agents at once: requests queue on the host and all complete.
$jobs = 1..5 | ForEach-Object { Start-Job -ScriptBlock { param($c) $o = & pwsh -NoProfile -File $c find -Days 2 -Limit 3 -Json 2>&1 | Out-String; [pscustomobject]@{ Exit = $LASTEXITCODE; Ok = [bool]($o | ConvertFrom-Json -ErrorAction SilentlyContinue) } } -ArgumentList $call }
$res = $jobs | Wait-Job -Timeout 180 | Receive-Job; $jobs | Remove-Job -Force
Check "5 concurrent clients all complete" (@($res | Where-Object { $_.Exit -eq 0 -and $_.Ok }).Count -eq 5) "ok=$(@($res | Where-Object { $_.Exit -eq 0 -and $_.Ok }).Count)"

# Folder cache across calls: a renamed folder's old path must not resolve.
$null = ViaHost @("folder-manage", "-Action", "Create", "-Folder", "Drafts", "-Name", "_suite-host")
$r1 = ViaHost @("find", "-Folder", "Drafts/_suite-host", "-Days", "1")
$null = ViaHost @("folder-manage", "-Action", "Rename", "-Folder", "Drafts/_suite-host", "-Name", "_suite-host2", "-Confirm")
$r2 = ViaHost @("find", "-Folder", "Drafts/_suite-host", "-Days", "1")
$r3 = ViaHost @("find", "-Folder", "Drafts/_suite-host2", "-Days", "1")
Check "host cache: old path fails after rename, new path works" ($r1.Exit -eq 0 -and $r2.Exit -eq 1 -and $r3.Exit -eq 0) "r1=$($r1.Exit) r2=$($r2.Exit) r3=$($r3.Exit)"
$null = ViaHost @("folder-manage", "-Action", "Delete", "-Folder", "Drafts/_suite-host2", "-Confirm")

# Benchmark: direct pwsh -File vs through the host.
$cases = @(
    @{ Name = "accounts-list"; Args = @("accounts-list") }
    @{ Name = "find 7d/50"; Args = @("find", "-Days", "7", "-Limit", "50") }
    @{ Name = "find 30d/200"; Args = @("find", "-Days", "30", "-Limit", "200") }
    @{ Name = "read 1d/10"; Args = @("read", "-Days", "1", "-Limit", "10") }
    @{ Name = "read-body"; Args = @("read-body", "-EntryID", $j.items[0].EntryID) }
    @{ Name = "folders -Under Inbox"; Args = @("folders", "-Under", "Inbox") }
)
$rows = foreach ($c in $cases) {
    $script = Join-Path $ScriptsDir ("outlook-" + $c.Args[0] + ".ps1")
    $rest = @($c.Args | Select-Object -Skip 1)
    $d = foreach ($i in 1..$Runs) { $sw = [Diagnostics.Stopwatch]::StartNew(); $null = & pwsh -NoProfile -File $script @rest 2>&1; $sw.ElapsedMilliseconds }
    $h = foreach ($i in 1..$Runs) { $sw = [Diagnostics.Stopwatch]::StartNew(); $null = & pwsh -NoProfile -File $call @($c.Args) 2>&1; $sw.ElapsedMilliseconds }
    [pscustomobject]@{ Case = $c.Name; DirectMs = ($d | Sort-Object)[[int]($Runs / 2)]; HostMs = ($h | Sort-Object)[[int]($Runs / 2)] }
}
$rows | Format-Table -AutoSize
Write-Host "(Both columns include starting the caller's own pwsh, ~200 ms; an agent already inside PowerShell skips that for host calls.)" -ForegroundColor DarkGray
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
